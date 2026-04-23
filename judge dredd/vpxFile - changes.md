Judge Dredd VBS — Optimization Changes
===============================================

1. EXPONENTIATION ELIMINATION: AudioFade, AudioPan, Pan (lines ~428-458)
   Replaced ^10 with chained multiply (t2=tmp*tmp, t4=t2*t2, t8=t4*t4, result=t8*t2).
   Handles both positive and negative branches separately to avoid sign issues.
   At ~100Hz per ball (RollingTimer + BallShadowUpdate + sound calls):
   with 5 balls, ~1,500 ^10 ops/sec eliminated.

2. PRE-COMPUTED INVERSE TABLE DIMENSIONS (line ~80, init ~136)
   Added InvTWHalf (2/table1.width) and InvTHHalf (2/table1.height) computed once in Table1_Init.
   AudioFade/AudioPan/Pan now multiply by inverse instead of dividing.
   Eliminates ~1,500 divisions/sec (same call frequency as #1).

3. BallVel ^2 ELIMINATION (line ~473)
   Replaced ball.VelX^2 + ball.VelY^2 with vx*vx + vy*vy.
   Also caches VelX/VelY into locals (2 COM reads instead of 2).
   BallVel is called from Vol, Pitch, RollingTimer, and collision handlers.
   At ~100Hz with 5 balls: ~500 ^2 ops/sec eliminated.

4. Vol ^2 ELIMINATION (line ~461)
   Replaced BallVel(ball)^2 with bv*bv using cached BallVel result.
   At ~100Hz with 5 balls: ~500 ^2 ops/sec eliminated.

5. VolZ ^2 ELIMINATION (line ~482)
   Replaced BallVelZ(ball)^2 with bvz*bvz.

6. OnBallBallCollision ^2 ELIMINATION (line ~590)
   Replaced velocity^2 with velocity*velocity.

7. PRE-BUILT ROLLING SOUND STRINGS (line ~78)
   Pre-built BallRollStr(0..15) = "fx_ballrolling0".."fx_ballrolling15" at module init.
   Pre-built BallDropStr(0..15) = "fx_ball_drop0".."fx_ball_drop15" at module init.
   Eliminates string concatenation ("fx_ballrolling" & b) inside RollingTimer_Timer.
   At ~100Hz with 5 balls: ~1,000 string allocs/sec eliminated.

8. RollingTimer_Timer INLINE REWRITE (line ~517)
   Inlined BallVel/Vol/Pitch computations directly in the loop body.
   Cached BOT(b).VelX/VelY/z/VelZ into locals (4 COM reads instead of 10+).
   Cached UBound(BOT) into local ubBot.
   Used pre-built string arrays for sound names.
   Nested If statements for ball drop sound (VBScript AND doesn't short-circuit).
   At ~100Hz with 5 balls: ~3,000 redundant COM reads/sec + ~1,000 function call overheads/sec eliminated.

9. BallShadowUpdate COM CACHING (line ~553)
   Cached BOT(b).X, .Y, .Z into locals bx, by, bz at top of each iteration.
   Each ball previously read X twice, Y twice, Z four times = 8 COM reads.
   Now: 3 COM reads per ball.
   Cached UBound(BOT) into local.
   Added .visible write guards (only write when value changes).
   At ~60Hz with 5 balls: ~1,500 COM reads/sec + ~300 redundant writes/sec eliminated.

10. DeadWorld_Timer COM CACHING (line ~914)
    Cached Nipple.RotY into local rotY at top of sub.
    Nipple.RotY was read ~15 times per tick. Now: 1 read + 1 write.
    At timer frequency (~100Hz): ~1,400 COM reads/sec eliminated.

11. Planet_Watch_Timer COM CACHING (line ~980)
    Cached Nipple.RotY into local. Was read 6 times per tick. Now: 1 read.
    At timer frequency: ~500 COM reads/sec eliminated.

12. Crane_X_Timer / Ball_Move_Timer COM CACHING (lines ~1327, ~1415)
    Cached Nipple.RotY into local in Case 3/4 blocks where it was read multiple times.

13. JDFlip_Timer FLIPPER ANGLE CACHING + DELTA GUARDS (line ~1453)
    Cached LeftFlipper.CurrentAngle, RightFlipper.CurrentAngle, etc. into locals.
    Added lastLFAngle/lastRFAngle/lastLF2Angle/lastRF2Angle tracking variables.
    Only writes RotZ/RotY when angle has actually changed.
    When flippers are at rest (majority of gameplay): eliminates 6 COM writes/tick.
    At timer frequency (~100Hz): ~600 COM writes/sec eliminated when flippers stable.


ESTIMATED TOTAL SAVINGS
========================
- Exponentiation: ~2,500 ^10 and ^2 ops/sec eliminated
- Division: ~1,500 divisions/sec eliminated
- COM reads: ~7,400 COM reads/sec eliminated
- COM writes: ~900 redundant writes/sec eliminated
- String allocs: ~1,000 string allocs/sec eliminated
- Function calls: ~1,000 function call overheads/sec eliminated

Conservative total: ~14,300 redundant operations/sec eliminated.
