Star Trek 25th Anniversary (Data East 1991) VBS — Optimization Changes
=======================================================================

1. FLIPPER TIMER 1ms -> 10ms (line ~1692)
   Changed RightFlipper.timerinterval from 1 to 10.
   100Hz is visually identical to 1000Hz for flipper correction.
   At 1000Hz: FlipperTricks x2 + FlipperNudge x2 = ~4000 calls/sec.
   At 100Hz: ~400 calls/sec. Saves ~3,600 function calls/sec.

2. SHARED GetBalls + PRESS GUARD IN FLIPPER TIMER (line ~1696)
   Added `If LFPress = 0 And RFPress = 0 Then Exit Sub` before
   FlipperNudge calls. FlipperNudge now accepts BOT array parameter
   instead of calling GetBalls internally.
   When flippers at rest: eliminates 100% of GetBalls + FlipperNudge.
   Saves ~200 GetBalls allocs/sec at 100Hz.

3. FlipperTricks — CACHED startangle/currentangle (line ~1870)
   Cached Flipper.startangle and Flipper.currentangle into locals
   sa/ca/absSa/absCa. Previously read 4-6 times per call.
   At 100Hz x 2 flippers: ~800 COM reads/sec eliminated.

4. FlipperNudge — CACHED currentangle (line ~1707)
   Cached Flipper1.currentangle into local ca1.
   Removed debug.print calls from hot path.
   At 100Hz x 2 flippers: ~200 COM reads/sec eliminated.

5. FlipperTrigger — CACHED ALL FLIPPER PROPERTIES (line ~1793)
   Cached Flipper.x/y/currentangle/Length into locals fx/fy/fca/flen.
   Inlined DistanceFromFlipper using cached values and PIover180.
   Previously 6-10 COM reads across FlipperTrigger+DistanceFromFlipper.
   At 100Hz with 4 balls: ~3,000 COM reads/sec eliminated.

6. EXPONENTIATION ELIMINATION — AudioFade/AudioPan ^10 (line ~2305)
   Replaced `tmp ^10` with chained multiply: t2=tmp*tmp, t4=t2*t2,
   t8=t4*t4, result=t8*t2.
   At 100Hz x 5 balls: ~1,000 ops/sec eliminated.

7. EXPONENTIATION ELIMINATION — BallVel ^2 (line ~2350)
   Replaced `ball.VelX ^2` + `ball.VelY ^2` with cached locals vx*vx + vy*vy.
   Eliminates 2 exponentiation ops + caches 2 COM reads per call.
   At 100Hz x 5 balls: ~1,000 ops/sec eliminated.

8. EXPONENTIATION ELIMINATION — Vol ^2, Volz ^2 (line ~2338)
   Replaced `BallVel(ball) ^2` with `bv * bv`.
   Replaced `ball.velz ^2` with cached `vz * vz`.

9. EXPONENTIATION ELIMINATION — VolPlayfieldRoll ^3 (line ~2354)
   Replaced `BallVel(ball) ^3` with cached `bv * bv * bv`.

10. EXPONENTIATION ELIMINATION — PitchPlayfieldRoll ^2 (line ~2359)
    Replaced `BallVel(ball) ^2 * 15` with cached `bv * bv * 15`.

11. PRE-COMPUTED INVERSE MULTIPLIERS (line ~16)
    Added InvTWHalf = 2/tablewidth and InvTHHalf = 2/tableheight.
    AudioPan/AudioFade multiply instead of divide.
    At 100Hz x 5 balls x 2 functions: ~1,000 divisions/sec eliminated.

12. PRE-COMPUTED TRIG CONSTANTS (line ~1738)
    Added PIover180, d180overPI at module level.
    Updated dSin, dCos, Radians, AnglePP (both instances).
    FlipperTrigger at 100Hz: ~800 divisions/sec eliminated.

13. Distance — ^2 to x*x (line ~1771)
    Replaced `(ax-bx)^2 + (ay-by)^2` with `dx*dx + dy*dy`.
    Called from FlipperTrigger and rubber dampener code.

14. BallSpeed — ^2 to x*x + COM caching (line ~1647)
    Cached VelX/VelY/VelZ into locals before computing.
    Used by CoRTracker.Update (~100Hz) and Dampener (~per-hit).

15. ROLLINGTIMER REWRITE — COM CACHING + INLINED AUDIO (line ~1005)
    Cached BOT(b).X/Y/Z/VelX/VelY/VelZ into locals per ball.
    Inlined BallVel, VolPlayfieldRoll, PitchPlayfieldRoll to avoid
    re-reading COM props. Pre-computed BS_d2/BS_d4/BS_d5/BS_d10.
    Original: ~15 COM reads per ball per frame.
    After: 6 COM reads per ball per frame.
    At 100Hz x 5 balls: ~4,500 COM reads/sec eliminated.

16. PRE-BUILT STRING ARRAYS — BallRollStr/RampLoopStr (line ~993)
    Pre-built "BallRoll_" & i and "RampLoop" & i at init.
    Eliminates 4 string concats per ball per frame (play + stop, both types).
    At 100Hz x 5 balls: ~2,000 string allocs/sec eliminated.

17. CACHED UBound(BOT) (multiple locations)
    UBound(BOT) cached into local ubBot in RollingTimer, DynamicBSUpdate.

18. FLIPPER SHADOW TIMER — DELTA GUARDS (line ~1075)
    FlipperTimer_Timer now caches currentangle and only writes
    RotZ/Roty when changed. Eliminates 4 COM writes/frame when
    flippers at rest (~80% of gameplay).
    At 60Hz: ~240 COM writes/sec eliminated.

19. DynamicBSUpdate REWRITE (line ~3080)
    - Cached BOT(s).X/Y/Z into bx/by/bz per ball.
    - Cached DSSources(iii)(0)/(1) into sx/sy per source.
    - Distance-squared gate before DistanceFast (skip when dsq > falloffSq).
    - Pre-computed invFalloff (multiply instead of divide).
    - Pre-computed DynBSFactor2/DynBSFactor3 (eliminates ^2/^3 per shadow).
    - Pre-computed TW_d2, BS_dAM, BS_d10, BS_d5, BS_d4, BS_d2.
    At 60Hz x 5 balls x 12 sources: ~3,000 COM reads/sec eliminated,
    ~50 Sqr/DistanceFast skipped, exponentiation eliminated.

20. LampTimer — CACHED chglamp 2D ARRAY (line ~546)
    Cached chglamp(ii,0) and chglamp(ii,1) into locals cIdx/cVal
    at top of loop. Eliminates redundant 2D array dereferences.
    During lamp storms: ~40-100 dereferences/tick eliminated.

21. CoRTracker — PRE-ALLOCATED ARRAYS (line ~2100)
    Changed Class_Initialize from redim ballvel(0) to redim ballvel(tnob).
    Eliminates per-frame ReDim checks in Update().


ESTIMATED TOTAL SAVINGS
========================
- Flipper timer reduction:     ~3,600 function calls/sec eliminated
- COM reads:                   ~10,000/sec eliminated (flipper, rolling, shadow)
- Exponentiation:              ~3,000/sec eliminated (^10, ^2, ^3)
- String allocs:               ~2,000/sec eliminated (BallRollStr, RampLoopStr)
- Division:                    ~2,000/sec eliminated (inverse multipliers, trig)
- COM writes:                  ~240/sec eliminated (flipper shadow guards)
- GetBalls allocs:             ~200/sec eliminated (shared flipper BOT)

Conservative total: ~21,000 redundant operations/sec eliminated.
