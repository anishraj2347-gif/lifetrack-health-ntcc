import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:lifetrack/src/ai/plane_a/critical_check.dart';
import 'package:lifetrack/src/model/health_record.dart';
import 'package:lifetrack/src/ui/theme.dart';
import 'package:lifetrack/src/ui/widgets/critical_prompt.dart';
import 'package:lifetrack/src/ui/widgets/parts.dart';

/// Every vibration goes through Haptics, one kind per kind of moment
/// (27 Nov 2026). These check what each moment plays.
void main() {
  late List<String> felt;

  /// Pumps [body] once, recording each vibration by its type.
  Future<void> pump(WidgetTester tester, Widget body, {bool still = false}) async {
    felt = [];
    tester.binding.defaultBinaryMessenger.setMockMethodCallHandler(SystemChannels.platform, (call) async {
      if (call.method == 'HapticFeedback.vibrate') felt.add((call.arguments as String?)?.replaceFirst('HapticFeedbackType.', '') ?? 'vibrate');
      return null;
    });
    addTearDown(() => tester.binding.defaultBinaryMessenger.setMockMethodCallHandler(SystemChannels.platform, null));
    await tester.pumpWidget(MaterialApp(
      theme: LT.theme(),
      builder: (context, child) => MediaQuery(data: MediaQuery.of(context).copyWith(disableAnimations: still), child: child!),
      home: Scaffold(body: body),
    ));
  }

  Widget button(String label, void Function(BuildContext) onPressed) => Builder(builder: (c) => TextButton(onPressed: () => onPressed(c), child: Text(label)));

  testWidgets('a card that opens gives a light tap; one that does not, nothing', (tester) async {
    var opened = 0;
    await pump(tester, ListView(children: [ShadCard(onTap: () => opened++, child: Text('Opens')), ShadCard(child: Text('Just reads'))]));
    await tester.pumpAndSettle();
    await tester.tap(find.text('Opens'));
    await tester.tap(find.text('Just reads'));
    await tester.pumpAndSettle();
    expect(opened, 1);
    expect(felt, ['lightImpact']);
  });

  testWidgets('the main button presses; the others tap', (tester) async {
    await pump(tester, Column(children: [SlabButton('Save', onPressed: () {}), SlabButton('Not now', variant: ButtonVariant.ghost, onPressed: () {})]));
    await tester.pumpAndSettle();
    await tester.tap(find.text('Save'));
    await tester.tap(find.text('Not now'));
    await tester.pumpAndSettle();
    expect(felt, ['mediumImpact', 'lightImpact']);
  });

  testWidgets('a switch ticks as it changes; a disabled one stays disabled', (tester) async {
    var on = false;
    await pump(tester, Column(children: [Switch(value: on, onChanged: withTick((v) => on = v)), Switch(value: false, onChanged: withTick<bool>(null))]));
    await tester.tap(find.byType(Switch).first);
    await tester.pumpAndSettle();
    expect(on, isTrue);
    expect(felt, ['selectionClick']);
    expect(tester.widget<Switch>(find.byType(Switch).last).onChanged, isNull);
  });

  testWidgets('an error under a form shakes once, with the error buzz', (tester) async {
    await pump(tester, Align(alignment: Alignment.topLeft, child: ErrorLine('Not saved.')));
    expect(felt, ['errorNotification']);
    await tester.pump(Duration(milliseconds: 40));
    final mid = tester.getTopLeft(find.text('Not saved.')).dx;
    await tester.pumpAndSettle();
    final rest = tester.getTopLeft(find.text('Not saved.')).dx;
    expect(rest, 20);
    expect(mid, isNot(rest), reason: 'it moves while it shakes');
    expect(felt, hasLength(1), reason: 'one buzz, not one per frame');
  });

  testWidgets('when the phone asks for less motion an error stays still, and still buzzes', (tester) async {
    await pump(tester, Align(alignment: Alignment.topLeft, child: ErrorLine('Not saved.')), still: true);
    await tester.pump(Duration(milliseconds: 40));
    expect(tester.getTopLeft(find.text('Not saved.')).dx, 20);
    expect(felt, ['errorNotification']);
    await tester.pumpAndSettle();
  });

  testWidgets('a problem in a snack bar buzzes', (tester) async {
    await pump(tester, button('Go', (c) => showProblem(ScaffoldMessenger.of(c), 'Could not save.')));
    await tester.tap(find.text('Go'));
    await tester.pumpAndSettle();
    expect(find.text('Could not save.'), findsOneWidget);
    expect(felt, ['errorNotification']);
    // Let the snack bar go by itself.
    await tester.pump(Duration(seconds: 5));
    await tester.pumpAndSettle();
  });

  testWidgets('the done moment buzzes success', (tester) async {
    await pump(tester, button('Go', (c) => celebrate(c, 'All done for today')));
    await tester.tap(find.text('Go'));
    await tester.pump();
    expect(felt, ['successNotification']);
    await tester.pumpAndSettle();
  });

  testWidgets('a reading at a danger level gives the strongest buzz; a value it cannot judge, a softer one', (tester) async {
    final at = DateTime(2026, 11, 27);
    HealthRecord report(LabResult r) => HealthRecord(id: 'r', patientId: 'p', kind: RecordKind.report, title: 'R', date: at, source: RecordSource.typedByPatient, results: [r]);
    for (final (result, want) in [
      (const LabResult(analyte: 'potassium', value: 6.5, unit: 'mmol/L'), 'vibrate'),
      (const LabResult(analyte: 'potassium', value: 65, unit: ''), 'warningNotification'),
    ]) {
      await pump(tester, button('Open', (c) => showCriticalPrompt(c, CriticalCheck.record(patient: null, at: at, record: report(result)))));
      await tester.tap(find.text('Open'));
      await tester.pumpAndSettle();
      expect(felt, [want], reason: '${result.value} ${result.unit}');
      await tester.tap(find.text('I have read this'));
      await tester.pumpAndSettle();
    }
  });
}
