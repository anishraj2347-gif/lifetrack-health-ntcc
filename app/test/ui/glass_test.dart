import 'dart:math' as math;

import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:lifetrack/src/ui/theme.dart';
import 'package:lifetrack/src/ui/widgets/parts.dart';

double _lum(Color c) {
  double ch(double v) => v <= 0.03928 ? v / 12.92 : math.pow((v + 0.055) / 1.055, 2.4).toDouble();
  return 0.2126 * ch(c.r) + 0.7152 * ch(c.g) + 0.0722 * ch(c.b);
}

double _ratio(Color a, Color b) {
  final (x, y) = (_lum(a), _lum(b));
  return (math.max(x, y) + 0.05) / (math.min(x, y) + 0.05);
}

/// The glass look (27 Nov 2026): a colour wash behind every page and
/// see-through cards over it. These keep it readable and cheap to draw.
void main() {
  tearDown(() => LT.dark = false);

  for (final dark in [false, true]) {
    final mode = dark ? 'dark' : 'light';

    test('$mode: the faintest text keeps 4.5:1 where the wash is strongest, on the page and on a card', () {
      LT.dark = dark;
      for (final (_, _, c) in LT.washes) {
        final painted = Color.alphaBlend(c.withValues(alpha: LT.washStrength), LT.ground);
        expect(_ratio(LT.ink3, painted), greaterThanOrEqualTo(4.5), reason: 'ink3 on $c');
      }
      expect(_ratio(LT.ink3, LT.washPeak), greaterThanOrEqualTo(4.5));
      expect(_ratio(LT.ink3, ShadCard.glassTone), greaterThanOrEqualTo(4.5));
    });

    test('$mode: washExtreme is the worst the wash can be behind glass', () {
      LT.dark = dark;
      for (final (_, _, c) in LT.washes) {
        if (dark) {
          expect(_lum(c), lessThanOrEqualTo(_lum(LT.washExtreme)), reason: '$c is lighter');
        } else {
          expect(_lum(c), greaterThanOrEqualTo(_lum(LT.washExtreme)), reason: '$c is darker');
        }
      }
    });

    test('$mode: the wash is never a health colour', () {
      LT.dark = dark;
      for (final (_, _, c) in LT.washes) {
        final hue = HSVColor.fromColor(c).hue;
        expect(hue < 15 || hue > 345, isFalse, reason: '$c is red');
        expect(hue > 70 && hue < 170, isFalse, reason: '$c is green or lime');
      }
    });

    testWidgets('$mode: a card is see-through glass with no live blur, its text judged on the worst wash', (tester) async {
      LT.dark = dark;
      await tester.pumpWidget(MaterialApp(
        theme: LT.theme(),
        home: Scaffold(body: Backdrop(child: ListView(children: [Text('On the page'), ShadCard(onTap: () {}, child: Text('On a card'))]))),
      ));
      await tester.pumpAndSettle();
      final card = find.byType(ShadCard);
      // The wash behind a card is already soft: a blur per card would cost a
      // GPU pass per card on every frame and look the same.
      expect(find.descendant(of: card, matching: find.byType(BackdropFilter)), findsNothing);
      final fill = tester
          .widgetList<DecoratedBox>(find.descendant(of: card, matching: find.byType(DecoratedBox)))
          .map((d) => d.decoration)
          .whereType<BoxDecoration>()
          .firstWhere((d) => d.gradient != null)
          .gradient!;
      expect(fill.colors.every((c) => c.a < 0.9), isTrue, reason: 'the wash shows through');
      GlassTone toneOf(String text) => tester.widget<GlassTone>(find.ancestor(of: find.text(text), matching: find.byType(GlassTone)).first);
      expect(toneOf('On a card').tone, ShadCard.glassTone);
      expect(toneOf('On the page').tone, LT.washPeak);
    });
  }
}
