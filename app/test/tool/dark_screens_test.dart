// Renders the five main tabs to PNGs for a visual check (and the report):
//   flutter test test/tool/dark_screens_test.dart --dart-define=SHOTS=<dir> [--dart-define=MODE=light]
// Skipped unless SHOTS is set.
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
import 'package:lifetrack/src/share/relay.dart';
import 'package:lifetrack/src/storage/byte_store.dart';
import 'package:lifetrack/src/ui/lifetrack_app.dart';
import 'package:lifetrack/src/ui/theme.dart';
import 'package:lifetrack/src/vault/backup_vault.dart';

import 'shot_fonts.dart';

const out = String.fromEnvironment('SHOTS');

void main() {
  testWidgets('dark screens', skip: out.isEmpty, (tester) async {
    await tester.runAsync(loadShotFonts);
    late AppServices s;
    await tester.runAsync(() async {
      s = await AppServices.open(
        bytes: InMemoryByteStore(), deviceKey: SecretKey(List.filled(32, 5)), remote: InMemoryRemoteVaultStore(),
        kdf: KdfParams.weakForTests(), cloudFactory: defaultCloud, clock: () => DateTime(2026, 10, 1, 10), demoTools: true,
        // Sharing switched on (in memory), so the Share tab shows what a real build shows.
        share: ShareConfig(relayUrl: 'http://127.0.0.1:1', registryKey: base64Encode((await (await Ed25519().newKeyPair()).extractPublicKey()).bytes)),
        relayOverride: (_) => InMemoryRelay(), idDirectoryOverride: () => InMemoryIdDirectory(),
      );
      await s.createProfile(name: 'Anjali Sharma', dob: DateTime(1984), sex: Sex.female);
      await s.addSampleRecords();
      await s.setThemeChoice(const String.fromEnvironment('MODE', defaultValue: 'dark'));
    });
    addTearDown(() {
      appearance.value = ThemeMode.system;
      LT.dark = false;
    });
    tester.view.physicalSize = const Size(1080, 2340);
    tester.view.devicePixelRatio = 2.6;
    addTearDown(tester.view.reset);
    final key = GlobalKey();
    await tester.pumpWidget(RepaintBoundary(key: key, child: LifeTrackApp(boot: () async => s)));
    await tester.pumpAndSettle();

    Future<void> shot(String name) async {
      await tester.pumpAndSettle();
      await tester.runAsync(() async {
        final img = await (key.currentContext!.findRenderObject()! as RenderRepaintBoundary).toImage(pixelRatio: 1);
        final png = await img.toByteData(format: ui.ImageByteFormat.png);
        File('$out/$name.png').writeAsBytesSync(png!.buffer.asUint8List());
      });
    }

    await shot('1-reminders');
    for (final (i, tab) in ['Records', 'Insights', 'Share', 'Me'].indexed) {
      await tester.tap(find.bySemanticsLabel(tab));
      await shot('${i + 2}-${tab.toLowerCase()}');
    }
  });
}
