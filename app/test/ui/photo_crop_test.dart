import 'dart:io';
import 'dart:ui' as ui;

import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:lifetrack/src/ui/widgets/add_person_parts.dart';
import 'package:lifetrack/src/ui/widgets/parts.dart';

/// Every crop of a bundled photo keeps its subject in view, and the opening
/// screens never stretch a photo past its own pixels (phone review, 27 Nov
/// 2026: the short Insights banner cut the faces off at the eyes, and the
/// lock screen's photo was stretched 1.3 times).
void main() {
  const photos = [Photos.family, Photos.report, Photos.doctor, Photos.pills, Photos.reading, Photos.parent];
  const photo = Size(1080, 807);

  /// The part of the photo in view when it covers [box] lined up by
  /// [align], as fractions of its width and height — worked out by the same
  /// calls the image painter makes.
  Rect shown(Size box, Alignment align) {
    final fit = applyBoxFit(BoxFit.cover, photo, box);
    final src = align.inscribe(fit.source, Offset.zero & photo);
    return Rect.fromLTRB(src.left / photo.width, src.top / photo.height, src.right / photo.width, src.bottom / photo.height);
  }

  testWidgets('every bundled photo is 1080 × 807, as Photos.aspect says', (tester) async {
    await tester.runAsync(() async {
      for (final p in photos) {
        final image = (await (await ui.instantiateImageCodec(File(p).readAsBytesSync())).getNextFrame()).image;
        expect((image.width, image.height), (1080, 807), reason: p);
      }
    });
    expect(Photos.aspect, photo.aspectRatio);
  });

  test('every crop the app makes keeps the subject in view, centred where the photo allows', () {
    for (final width in [320.0, 360.0, 384.0, 411.0, 480.0, 501.0, 600.0]) {
      final boxes = [
        for (final h in [112.0, 150.0, 160.0, 170.0, 180.0]) Size(width - 32, h), // banners
        Size(width, 240), // the add-person pages
        for (final screen in [640.0, 780.0, 891.0, 1085.0]) Size(width, PhotoHero.photoHeight(Size(width, screen))),
      ];
      for (final p in photos) {
        final at = Photos.focus[p]!;
        for (final box in boxes) {
          final r = shown(box, Photos.align(p, box));
          const e = 1e-9;
          expect(r.left > -e && r.top > -e && r.right < 1 + e && r.bottom < 1 + e, isTrue, reason: '$p in $box stays inside the photo');
          expect(r.contains(at), isTrue, reason: '$p in $box shows its subject ($r)');
          // Centred on the subject, unless that would run past an edge.
          expect((r.center.dx - at.dx).abs() < 1e-6 || r.left < e || r.right > 1 - e, isTrue, reason: '$p in $box: centred across');
          expect((r.center.dy - at.dy).abs() < 1e-6 || r.top < e || r.bottom > 1 - e, isTrue, reason: '$p in $box: centred down');
        }
      }
    }
  });

  test('the Insights banner on the phone of the review shows both faces, chin to hairline', () {
    // 501 points wide (1080 pixels at 345 dpi): the banner is 469 × 112.
    const box = Size(469, 112);
    final r = shown(box, Photos.align(Photos.reading, box));
    expect(r.top, lessThanOrEqualTo(0.25), reason: 'the hairline, at a quarter of the photo');
    expect(r.bottom, greaterThanOrEqualTo(0.52), reason: 'both chins');
    final centred = shown(box, Alignment.center);
    expect(centred.top, greaterThan(0.33), reason: 'the old centred crop cut them at the eyes');
  });

  test('the opening photo is never taller than the photo is at the screen width', () {
    for (final width in [320.0, 360.0, 411.0, 501.0, 600.0, 800.0]) {
      for (final height in [560.0, 640.0, 780.0, 891.0, 1085.0, 1280.0]) {
        final h = PhotoHero.photoHeight(Size(width, height));
        expect(h, lessThanOrEqualTo(width / Photos.aspect + 1e-9), reason: '$width × $height');
        expect(h, lessThanOrEqualTo(480));
      }
    }
  });

  group('on screen', () {
    Future<void> phone(WidgetTester tester) async {
      tester.view.physicalSize = const Size(1080, 2340);
      tester.view.devicePixelRatio = 2.15625; // 501 × 1085 points
      addTearDown(tester.view.reset);
    }

    testWidgets('a banner lines up on its photo\'s subject', (tester) async {
      await phone(tester);
      await tester.pumpWidget(MaterialApp(home: Scaffold(body: ListView(children: [PhotoBanner(Photos.reading, height: 112)]))));
      await tester.pumpAndSettle();
      final image = tester.widget<Image>(find.byType(Image));
      final size = tester.getSize(find.byType(Image));
      expect(size.width, closeTo(469, 0.5));
      expect(size.height, 112);
      expect(image.alignment, Photos.align(Photos.reading, size));
      expect((image.alignment as Alignment).y, lessThan(0), reason: 'the faces are above the middle of the photo');
    });

    testWidgets('the add-person photo lines up on its subject', (tester) async {
      await phone(tester);
      await tester.pumpWidget(MaterialApp(home: AddPersonScaffold(photo: Photos.parent, eyebrow: 'Family', title: 'Who are you adding?', children: [])));
      await tester.pumpAndSettle();
      final image = tester.widget<Image>(find.byType(Image));
      expect(image.alignment, Photos.align(Photos.parent, tester.getSize(find.byType(Image))));
    });

    testWidgets('the opening photo is drawn no taller than it is, its words at the foot of the screen', (tester) async {
      await phone(tester);
      await tester.pumpWidget(MaterialApp(home: Scaffold(body: PhotoHero(Photos.reading, children: [Text('LifeTrack is locked')]))));
      await tester.pumpAndSettle();
      final size = tester.getSize(find.byType(Image));
      expect(size.width, closeTo(501, 0.5));
      expect(size.height, lessThanOrEqualTo(size.width / Photos.aspect + 1e-6));
      expect(tester.getBottomLeft(find.text('LifeTrack is locked')).dy, greaterThan(900), reason: 'in reach, not floating under the photo');
    });
  });
}
