// Renders the screens the NTCC report shows, from the current code, to PNGs:
//   flutter test test/tool/report_screens_test.dart --dart-define=SHOTS=<dir>
// Skipped unless SHOTS is set. A common phone size (1080 × 2340 pixels at
// 2.625, 411 points wide), light mode unless a file name says dark; the
// ten-person sample family with its photos. The welcome, first step and
// lock come from photo_screens_test.dart (MODE=light, DPR=2.625).
import 'dart:convert';
import 'dart:io';
import 'dart:ui' as ui;

import 'package:cryptography/cryptography.dart';
import 'package:flutter/material.dart';
import 'package:flutter/rendering.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:lifetrack/src/ai/plane_a/critical_check.dart';
import 'package:lifetrack/src/app/app_services.dart';
import 'package:lifetrack/src/crypto/key_derivation.dart';
import 'package:lifetrack/src/identity/lifetrack_ids.dart';
import 'package:lifetrack/src/model/health_record.dart';
import 'package:lifetrack/src/model/patient.dart';
import 'package:lifetrack/src/share/relay.dart';
import 'package:lifetrack/src/storage/byte_store.dart';
import 'package:lifetrack/src/ui/lifetrack_app.dart';
import 'package:lifetrack/src/ui/screens/add_family_screen.dart';
import 'package:lifetrack/src/ui/screens/add_record_screen.dart';
import 'package:lifetrack/src/ui/screens/clinical_profile_screen.dart';
import 'package:lifetrack/src/ui/screens/emergency_screen.dart';
import 'package:lifetrack/src/ui/screens/family_overview_screen.dart';
import 'package:lifetrack/src/ui/screens/record_screen.dart';
import 'package:lifetrack/src/ui/screens/scan_report_screen.dart';
import 'package:lifetrack/src/ui/screens/test_trend_screen.dart';
import 'package:lifetrack/src/ui/theme.dart';
import 'package:lifetrack/src/ui/widgets/critical_prompt.dart';
import 'package:lifetrack/src/ui/widgets/person_avatar.dart';
import 'package:lifetrack/src/vault/backup_vault.dart';

import 'shot_fonts.dart';

const out = String.fromEnvironment('SHOTS');
const dpr = 2.625;

void main() {
  testWidgets('report screens', skip: out.isEmpty, (tester) async {
    await tester.runAsync(loadShotFonts);
    late AppServices s;
    await tester.runAsync(() async {
      s = await AppServices.open(
        bytes: InMemoryByteStore(), deviceKey: SecretKey(List.filled(32, 5)), remote: InMemoryRemoteVaultStore(),
        kdf: KdfParams.weakForTests(), cloudFactory: defaultCloud, clock: () => DateTime(2026, 10, 1, 10), demoTools: true,
        share: ShareConfig(relayUrl: 'http://127.0.0.1:1', registryKey: base64Encode((await (await Ed25519().newKeyPair()).extractPublicKey()).bytes)),
        relayOverride: (_) => InMemoryRelay(), idDirectoryOverride: () => InMemoryIdDirectory(),
      );
      await s.createProfile(name: 'Anjali Sharma', dob: DateTime(1984), sex: Sex.female, demo: true);
      await s.addDemoFamily(photo: bundledPhoto);
      await s.setThemeChoice('light');
    });
    tester.view.physicalSize = const Size(1080, 2340);
    tester.view.devicePixelRatio = dpr;
    tester.view.padding = const FakeViewPadding(top: 63, bottom: 42); // status and gesture bars
    addTearDown(tester.view.reset);
    addTearDown(() {
      appearance.value = ThemeMode.system;
      LT.dark = false;
    });
    final key = GlobalKey();
    await tester.pumpWidget(RepaintBoundary(key: key, child: LifeTrackApp(boot: () async => s)));
    await tester.pumpAndSettle();

    // Pumps long enough for entrances (the photo drift is 1.4 s) without
    // waiting on a repeating animation, as pumpAndSettle would.
    Future<void> settle() async {
      for (var i = 0; i < 20; i++) {
        await tester.pump(const Duration(milliseconds: 100));
      }
    }

    Future<void> shot(String name) async {
      await settle();
      // A test does not wait for pictures: decode each one first, through
      // the provider the screen asks for (photos, faces, report pages).
      await tester.runAsync(() async {
        for (final e in find.byType(Image).evaluate()) {
          await precacheImage((e.widget as Image).image, e, onError: (_, _) {});
        }
        for (final e in find.byWidgetPredicate((w) => w is DecoratedBox && w.decoration is BoxDecoration && (w.decoration as BoxDecoration).image != null).evaluate()) {
          await precacheImage(((e.widget as DecoratedBox).decoration as BoxDecoration).image!.image, e, onError: (_, _) {});
        }
      });
      await settle();
      await tester.runAsync(() async {
        final img = await (key.currentContext!.findRenderObject()! as RenderRepaintBoundary).toImage(pixelRatio: dpr);
        final png = await img.toByteData(format: ui.ImageByteFormat.png);
        File('$out/$name.png').writeAsBytesSync(png!.buffer.asUint8List());
      });
    }

    final nav = tester.state<NavigatorState>(find.byType(Navigator).first);
    Future<void> page(Widget screen, String name, {Future<void> Function()? then}) async {
      nav.push(MaterialPageRoute<void>(builder: (_) => screen));
      await settle();
      await shot(name);
      if (then != null) await then();
      nav.pop();
      await settle();
    }

    // The five tabs.
    await shot('1-reminders');
    await tester.tap(find.bySemanticsLabel('Records'));
    await shot('2-records');
    await tester.tap(find.bySemanticsLabel('Insights'));
    await settle();
    // The button sits below the fold: bring it to the middle, clear of the bars.
    await tester.runAsync(() => Scrollable.ensureVisible(tester.element(find.text('Turn on for 24 hours')), alignment: 0.5));
    await settle();
    await tester.tap(find.text('Turn on for 24 hours'));
    await settle();
    await tester.dragFrom(const Offset(6, 400), const Offset(0, 3000)); // back to the top
    await shot('3-insights');
    await tester.tap(find.bySemanticsLabel('Share'));
    await shot('4-share');
    await tester.tap(find.bySemanticsLabel('Me'));
    await shot('5-me');

    // Records: a report with its original, a test over time, adding and scanning.
    final me = s.selfPatientId!;
    final lipid = s.repository.recordsOf(me).firstWhere((r) => r.kind == RecordKind.report && r.title == 'Lipid panel');
    await page(RecordScreen(recordId: lipid.id), 'record');
    await page(const TestTrendScreen(analyte: 'hba1c', unit: '%'), 'trend');
    await page(const AddRecordScreen(), 'add-record');
    final original = await tester.runAsync(() => s.originals(lipid));
    await page(ScanReportScreen(incoming: [for (final b in original ?? const []) (bytes: b, mime: b[0] == 0x89 ? 'image/png' : 'image/jpeg')]), 'scan');

    // Safety: allergies, the Emergency ID and its QR, the urgent prompt.
    await page(const ClinicalProfileScreen(), 'allergies');
    await page(const EmergencyScreen(), 'emergency', then: () async {
      await tester.tap(find.text('Penicillin')); // anywhere on the card turns it over
      await shot('emergency-qr');
    });
    final at = DateTime(2026, 10, 1, 10);
    final urgent = HealthRecord(id: 'u', patientId: me, kind: RecordKind.report, title: 'Electrolytes', date: at, source: RecordSource.typedByPatient,
        results: const [LabResult(analyte: 'potassium', value: 6.5, unit: 'mmol/L')]);
    showCriticalPrompt(nav.context, CriticalCheck.record(patient: s.repository.patient(me), at: at, record: urgent));
    await shot('urgent');
    await tester.tap(find.text('I have read this'));
    await settle();

    // Family.
    await page(const FamilyOverviewScreen(), 'family-tree');
    await page(const AddFamilyScreen(), 'add-person');

    // Dark mode.
    await tester.runAsync(() => s.setThemeChoice('dark'));
    appearance.value = ThemeMode.dark; // as Me → Appearance does
    await tester.tap(find.bySemanticsLabel('Reminders'));
    await shot('dark-reminders');
    await tester.tap(find.bySemanticsLabel('Records'));
    await shot('dark-records');
  });
}
