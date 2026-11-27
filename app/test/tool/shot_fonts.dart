import 'dart:io';

import 'package:flutter/services.dart';

/// Loads the app's text and icon fonts, so a render in a test draws real
/// letters and icons instead of boxes. Call inside `tester.runAsync`.
Future<void> loadShotFonts() async {
  for (final family in ['Poppins', 'Sora']) {
    final loader = FontLoader(family);
    for (final f in Directory('fonts').listSync().whereType<File>().where((f) => f.path.endsWith('.ttf') && f.uri.pathSegments.last.startsWith('$family-'))) {
      loader.addFont(Future.value(ByteData.sublistView(f.readAsBytesSync())));
    }
    await loader.load();
  }
  // Icon fonts, so icons are drawn instead of empty boxes.
  const lucide = 'packages/lucide_icons_flutter';
  for (final (family, asset) in [
    ('Lucide', 'assets/lucide.ttf'),
    for (final w in [100, 200, 300, 400, 500, 600]) ('Lucide$w', 'assets/build_font/LucideVariable-w$w.ttf'),
  ]) {
    await (FontLoader('$lucide/$family')..addFont(rootBundle.load('$lucide/$asset'))).load();
  }
  await (FontLoader('MaterialIcons')..addFont(rootBundle.load('fonts/MaterialIcons-Regular.otf'))).load();
}
