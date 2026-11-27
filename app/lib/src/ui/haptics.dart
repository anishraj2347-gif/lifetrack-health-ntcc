import 'package:flutter/services.dart';

/// Every vibration LifeTrack makes, named by what the moment means, so each
/// kind of moment always feels the same (27 Nov 2026). Call these, never
/// [HapticFeedback] directly. On Android they play through the screen's own
/// touch feedback: they follow the phone's vibration setting and need no
/// VIBRATE permission (the release manifest asks for none).
abstract final class Haptics {
  /// A choice moves: a tab, a chip, a segment, a switch, a tick box.
  static void select() => HapticFeedback.selectionClick();

  /// Something opens or is pressed: a card, a row, a round or outline button.
  static void tap() => HapticFeedback.lightImpact();

  /// A page's main action: a filled button, a swipe that ticks a dose.
  static void press() => HapticFeedback.mediumImpact();

  /// It worked: the "done" moment.
  static void success() => HapticFeedback.successNotification();

  /// Please read this: findings shown after a save that are not urgent.
  static void warning() => HapticFeedback.warningNotification();

  /// It did not work: an error under a form, or in a snack bar.
  static void error() => HapticFeedback.errorNotification();

  /// A reading at a danger level: the strongest buzz the app makes.
  static void alert() => HapticFeedback.vibrate();
}

/// A switch's or tick box's [onChanged] that also ticks under the finger.
/// Null stays null, so a disabled control stays disabled.
ValueChanged<T>? withTick<T>(ValueChanged<T>? onChanged) => onChanged == null
    ? null
    : (v) {
        Haptics.select();
        onChanged(v);
      };
