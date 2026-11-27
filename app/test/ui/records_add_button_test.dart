import 'package:cryptography/cryptography.dart';
import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:lifetrack/src/app/app_services.dart';
import 'package:lifetrack/src/crypto/key_derivation.dart';
import 'package:lifetrack/src/model/patient.dart';
import 'package:lifetrack/src/storage/byte_store.dart';
import 'package:lifetrack/src/ui/lifetrack_app.dart';
import 'package:lifetrack/src/ui/screens/add_record_screen.dart';
import 'package:lifetrack/src/ui/widgets/parts.dart';
import 'package:lifetrack/src/vault/backup_vault.dart';

/// Records' add button (phone test, 27 Nov 2026): the floating one sat over a
/// card's title at the top of the page. At the top the masthead holds it; the
/// floating one comes once that has scrolled away, and steps aside while you
/// scroll down.
void main() {
  testWidgets('at the top the add button is in the masthead and nothing floats over the records', (tester) async {
    final handle = tester.ensureSemantics();
    late AppServices s;
    await tester.runAsync(() async {
      s = await AppServices.open(
        bytes: InMemoryByteStore(), deviceKey: SecretKey(List.filled(32, 5)), remote: InMemoryRemoteVaultStore(),
        kdf: KdfParams.weakForTests(), cloudFactory: defaultCloud, clock: () => DateTime(2026, 10, 1, 10),
      );
      await s.createProfile(name: 'Anjali Sharma', dob: DateTime(1984), sex: Sex.female);
      await s.addSampleRecords();
    });
    tester.view.physicalSize = const Size(1080, 2340);
    tester.view.devicePixelRatio = 2.6;
    addTearDown(tester.view.reset);
    await tester.pumpWidget(LifeTrackApp(boot: () async => s));
    await tester.pumpAndSettle();
    await tester.tap(find.bySemanticsLabel('Records'));
    await tester.pumpAndSettle();

    bool floating() => !tester.widget<GlassFab>(find.byType(GlassFab)).hidden;
    final add = find.widgetWithText(SlabButton, 'Add');
    expect(add, findsOneWidget);
    expect(floating(), isFalse, reason: 'at the top nothing floats over the records');

    // Drags start in the list's left margin, clear of any card's own gestures.
    await tester.dragFrom(const Offset(6, 650), const Offset(0, -600));
    await tester.pumpAndSettle();
    expect(floating(), isFalse, reason: 'it steps aside while you scroll down');
    await tester.dragFrom(const Offset(6, 300), const Offset(0, 150));
    await tester.pumpAndSettle();
    expect(floating(), isTrue, reason: 'scrolling back up mid-list brings it');
    await tester.dragFrom(const Offset(6, 300), const Offset(0, 2000));
    await tester.pumpAndSettle();
    expect(floating(), isFalse, reason: 'back at the top the masthead has it');

    await tester.tap(add);
    await tester.pumpAndSettle();
    expect(find.byType(AddRecordScreen), findsOneWidget);
    handle.dispose();
  });
}
