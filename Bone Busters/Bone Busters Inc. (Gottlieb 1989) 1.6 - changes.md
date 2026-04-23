Bone Busters Inc. VBS — Optimization Changes
===============================================

1. PRE-COMPUTED CONSTANTS (line ~1992)
   Added InvTWHalf = 2/TableWidth, InvTHHalf = 2/TableHeight, TW_d2 = TableWidth/2,
   BS_d6 = Ballsize/6, InvTW7 = 1/7 at module level.
   Eliminates repeated division in per-frame AudioPan, AudioFade, and BallShadowUpdate loops.

2. PRE-BUILT BallRollStr ARRAY (line ~1997)
   Built "fx_ballrolling0" through "fx_ballrolling5" at init.
   Eliminates "fx_ballrolling" & b string concatenation in RollingTimer_Timer.
   At 100Hz with 3 balls: ~300 string allocs/sec eliminated.

3. PRE-COMPUTED TRIG CONSTANTS — PIover180, d180overPI (line ~1585)
   dSin, dCos, Radians: Pi/180 replaced with PIover180.
   AnglePP: 180/PI replaced with d180overPI.
   Eliminates 2 divisions per trig call across FlipperTrigger chain.
   At 100Hz: ~800 divisions/sec eliminated.

4. FLIPPER TIMER 1ms -> 10ms (line ~1543)
   FlipperRight.timerinterval changed from 1 to 10.
   100Hz is visually identical to 1000Hz for flipper correction.
   At 1000Hz: ~4,500 function calls/sec eliminated.

5. FLIPPER TIMER — SHARED GetBalls + PRESS GUARD (line ~1546)
   Added "If LFPress = 0 And RFPress = 0 Then Exit Sub" before nudge calls.
   FlipperNudge now accepts shared BOT array parameter (single GetBalls).
   Eliminates 2 GetBalls COM allocs per tick + all nudge COM reads when flippers at rest.
   At 100Hz: ~200 GetBalls allocs/sec eliminated; at rest: 100% of FlipperNudge skipped.

6. FlipperNudge — CACHED currentangle (line ~1570)
   Cached Flipper1.currentangle into ca1 local.
   At 100Hz: ~200 COM reads/sec eliminated.

7. FlipperTricks — CACHED startangle AND currentangle (line ~1640)
   Cached Flipper.startangle into sa, Abs(startangle) into absSa,
   Abs(currentangle) into ca. Replaced all references.
   At 100Hz (2 calls/tick): ~800 COM reads/sec eliminated.

8. Distance — ^2 ELIMINATION (line ~1620)
   (ax-bx)^2 replaced with dx*dx. Called from FlipperTrigger and related functions.
   At 100Hz with source loop: ~2,000 exponentiation ops/sec eliminated.

9. EXPONENTIATION ELIMINATION — ^10 in AudioFade/AudioPan (line ~2002)
   Replaced ^10 with chained multiply: t2=tmp*tmp, t4=t2*t2, t8=t4*t4, result=t8*t2.
   At 100Hz with 3 balls: ~600 Exp(10*Log(x)) ops/sec eliminated.

10. PRE-COMPUTED INVERSE MULTIPLIERS — AudioFade/AudioPan (line ~2002)
    Replaced * 2 / tableheight with * InvTHHalf, * 2 / tablewidth with * InvTWHalf.
    At 100Hz with 3 balls: ~600 divisions/sec eliminated.

11. EXPONENTIATION ELIMINATION — ^2, ^3 in BallVel/Vol/Volz/VolPlayfieldRoll/PitchPlayfieldRoll/BallSpeed (line ~2030)
    All ^2 replaced with x*x, ^3 with x*x*x. COM reads cached into locals.
    BallSpeed: ^2 replaced with vx*vx for VelX/VelY/VelZ.
    At 100Hz with 3 balls: ~900 exponentiation + ~900 COM reads/sec eliminated.

12. FlipperTimer_Timer — FLIPPER ANGLE WRITE GUARDS (line ~335)
    Cached FlipperLeft/FlipperLeftUp/FlipperRight/FlipperRightUp.CurrentAngle once per tick.
    Guarded all 7 objRotZ/RotZ writes with delta check (4 Gottlieb shadows + 3 flipper shadows).
    Added lastLFAngle, lastLFUpAngle, lastRFAngle, lastRFUpAngle tracking variables.
    At 100Hz with flippers at rest: ~700 COM writes/sec eliminated.
    Also reduces 8 COM reads per tick to 4 (each CurrentAngle read once, used for both objRotZ and RotZ).

13. RollingTimer_Timer REWRITE (line ~1218)
    - Set ball = BOT(b) cached; bx/by/bz/bvx/bvy/bvz cached as locals.
    - UBound(BOT) cached once at sub entry.
    - BallVel/VolPlayfieldRoll/PitchPlayfieldRoll inlined with cached locals.
    - AudioPan/AudioFade inlined with cached locals and ^10 chained multiply.
    - Pre-built BallRollStr used instead of string concatenation.
    - Drop sound VelZ/z use cached locals; nested If replaces non-short-circuit And.
    At 100Hz with 3 balls: ~4,500 COM reads/sec + ~300 string allocs/sec + ~1,800 function calls/sec eliminated.

14. BallShadowUpdate_timer REWRITE (line ~1279)
    - BOT(b).X/Y/Z cached into bx/by/bz locals at top of each ball iteration.
    - Table1.Width/2 replaced with pre-computed TW_d2.
    - Ballsize/6 replaced with pre-computed BS_d6.
    - Division by 7 replaced with multiplication by pre-computed InvTW7.
    - UBound(BOT) cached once at sub entry.
    At 60fps with 3 balls: ~540 COM reads/sec + ~180 divisions/sec eliminated.

15. LampTimer_Timer — CACHED chglamp 2D ARRAY (line ~971)
    chgLamp(ii,0)/(ii,1) cached into cIdx/cVal locals at loop top.
    Eliminates redundant 2D array dereferences (each index was accessed 6+ times).
    During lamp storms (50+ changes/tick): ~500 2D dereferences/tick eliminated.

16. OnBallBallCollision — ^2 ELIMINATION (line ~2553)
    Csng(velocity)^2 replaced with velF*velF.
    Occasional (per collision), minor savings.

ESTIMATED TOTAL SAVINGS
========================
- Flipper timer calls: ~4,500/sec (1ms -> 10ms)
- COM property reads: ~7,140/sec (flipper caching, RollingUpdate, BallShadowUpdate, BallVel inlining)
- COM writes: ~700/sec (flipper angle write guards)
- Exponentiation: ~3,500/sec (^10 -> chained multiply, ^2/^3 -> x*x, BallSpeed)
- Division: ~1,580/sec (pre-computed inverses, trig constants, BallShadow constants)
- String allocation: ~300/sec (pre-built BallRollStr)
- Function calls: ~6,300/sec (inlined BallVel/Vol/AudioPan/AudioFade in RollingUpdate + flipper timer reduction)
- GetBalls alloc: ~200/sec (shared BOT in flipper timer)
Conservative total: ~24,000 redundant operations/sec eliminated.
