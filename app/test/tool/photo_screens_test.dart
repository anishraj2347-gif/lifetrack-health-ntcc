// Renders every screen that shows one of the bundled photos to PNGs, at a
// phone's size and pixel density, to check how each photo is cropped and how
// sharp it is:
//   flutter test test/tool/photo_screens_test.dart --dart-define=SHOTS=<dir>
//       [--dart-define=MODE=light] [--dart-define=DPR=3]
// Skipped unless SHOTS is set. The default is the Galaxy S25 of the phone
// tests: 1080 × 2340 pixels with the display zoomed out (345 dpi).
import 'dart:convert';
import 'dart:io';
import 'dart:ui' as ui;

import 'package:cryptography/cryptography.dart';
import 'package:flutter/material.dart';
import 'package:flutter/rendering.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:lifetrack/src/app/app_services.dart';
import 'package:lifetrack/src/crypto/key_derivation.dart';
import 'package:lifetrack/src/identity/lifetrack_ids.dart';
import 'package:lifetrack/src/model/patient.dart';
import 'package:lifetrack/src/security/unlocker.dart';
import 'package:lifetrack/src/share/relay.dart';
import 'package:lifetrack/src/storage/byte_store.dart';
import 'package:lifetrack/src/ui/lifetrack_app.dart';
import 'package:lifetrack/src/ui/screens/add_family_screen.dart';
import 'package:lifetrack/src/ui/screens/add_record_screen.dart';
import 'package:lifetrack/src/ui/theme.dart';
import 'package:lifetrack/src/vault/backup_vault.dart';

import 'shot_fonts.dart';

const out = String.fromEnvironment('SHOTS');
const mode = String.fromEnvironment('MODE', defaultValue: 'dark');
final dpr = double.parse(const String.fromEnvironment('DPR', defaultValue: '2.15625'));

/// An app lock that is on and never opens.
class _Locked implements Unlocker {
  @override
  Future<bool> available() async => true;
  @override
  Future<bool> unlock(String reason) async => false;
}

void main() {
  final key = GlobalKey();

  Future<AppServices> open({bool profile = true}) async {
    final s = await AppServices.open(
      bytes: InMemoryByteStore(), deviceKey: SecretKey(List.filled(32, 5)), remote: InMemoryRemoteVaultStore(),
      kdf: KdfParams.weakForTests(), cloudFactory: defaultCloud, clock: () => DateTime(2026, 10, 1, 10), demoTools: true,
      // Sharing switched on (in memory), so the Share tab shows what a real build shows.
      share: ShareConfig(relayUrl: 'http://127.0.0.1:1', registryKey: base64Encode((await (await Ed25519().newKeyPair()).extractPublicKey()).bytes)),
      relayOverride: (_) => InMemoryRelay(), idDirectoryOverride: () => InMemoryIdDirectory(),
    );
    if (profile) await s.createProfile(name: 'Anjali Sharma', dob: DateTime(1984), sex: Sex.female);
    await s.setThemeChoice(mode);
    return s;
  }

  Future<void> start(WidgetTester tester, {bool profile = true, Unlocker? unlocker}) async {
    await tester.runAsync(loadShotFonts);
    late AppServices s;
    await tester.runAsync(() async => s = await open(profile: profile));
    tester.view.physicalSize = const Size(1080, 2340);
    tester.view.devicePixelRatio = dpr;
    tester.view.padding = const FakeViewPadding(top: 100, bottom: 48); // status and gesture bars
    addTearDown(tester.view.reset);
    addTearDown(() {
      appearance.value = ThemeMode.system;
      LT.dark = false;
    });
    await tester.pumpWidget(RepaintBoundary(key: key, child: LifeTrackApp(boot: () async => s, unlocker: unlocker)));
    await tester.pumpAndSettle();
  }

  Future<void> shot(WidgetTester tester, String name) async {
    // A test does not wait for pictures to decode: decode each one first,
    // through the same provider (and size) the screen asks for.
    await tester.runAsync(() async {
      for (final e in find.byType(Image).evaluate()) {
        await precacheImage((e.widget as Image).image, e, onError: (_, _) {});
      }
    });
    await tester.pumpAndSettle();
    await tester.runAsync(() async {
      final img = await (key.currentContext!.findRenderObject()! as RenderRepaintBoundary).toImage(pixelRatio: dpr);
      final png = await img.toByteData(format: ui.ImageByteFormat.png);
      File('$out/$name.png').writeAsBytesSync(png!.buffer.asUint8List());
    });
  }

  testWidgets('first run: the welcome and the first step', skip: out.isEmpty, (tester) async {
    await start(tester, profile: false);
    await shot(tester, '1-welcome');
    await tester.tap(find.text('Get started'));
    await tester.pumpAndSettle();
    await shot(tester, '2-first-step');
  });

  testWidgets('the lock', skip: out.isEmpty, (tester) async {
    await start(tester, unlocker: _Locked());
    expect(find.text('LifeTrack is locked'), findsOneWidget);
    await shot(tester, '3-lock');
  });

  testWidgets('a new profile: every empty page has a photo', skip: out.isEmpty, (tester) async {
    await start(tester);
    await shot(tester, '4-reminders');
    for (final (i, tab) in ['Records', 'Insights', 'Share'].indexed) {
      await tester.tap(find.bySemanticsLabel(tab));
      await tester.pumpAndSettle();
      await shot(tester, '${i + 5}-${tab.toLowerCase()}');
    }
    final nav = tester.state<NavigatorState>(find.byType(Navigator).first);
    nav.push(MaterialPageRoute<void>(builder: (_) => const AddRecordScreen()));
    await tester.pumpAndSettle();
    await shot(tester, '8-add-record');
    nav.pop();
    await tester.pumpAndSettle();
    nav.push(MaterialPageRoute<void>(builder: (_) => const AddFamilyScreen()));
    await tester.pumpAndSettle();
    await shot(tester, '9-add-family');
    await tester.tap(find.text('An adult who wants me to manage'));
    await tester.pumpAndSettle();
    await shot(tester, '10-add-adult');
  });
}
