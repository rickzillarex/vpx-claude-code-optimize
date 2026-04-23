Indiana Jones TPA VBS — Optimization Changes
===============================================

1. SHARED GetBalls IN RealTimeUpdates (line ~1015)
   RealTimeUpdates (MotorCallback, runs every frame at 60Hz) previously called
   RollingSoundUpdate and BallShadowUpdate, each calling GetBalls independently.
   Now calls GetBalls once and passes the BOT array to new BOT-accepting variants
   (RollingSoundUpdateBOT, BallShadowUpdateBOT). Original subs preserved as wrappers.
   At 60Hz: 60 GetBalls COM calls/sec eliminated.

2. EXPONENTIATION ELIMINATION: Pan, AudioFade (lines ~873-902)
   Replaced ^10 with chained multiply (t2=tmp*tmp, t4=t2*t2, t8=t4*t4, result=t8*t2).
   Handles positive/negative branches separately.
   Pan and AudioFade are called from RollingSoundUpdateBOT (per ball per frame),
   PlaySoundAt, PlaySoundAtBall, PlaySoundAtBallVol, OnBallBallCollision, etc.
   At 60Hz with 7 balls: ~840 ^10 ops/sec eliminated from rolling alone.

3. PRE-COMPUTED INVERSE TABLE DIMENSIONS (line ~65, init ~158)
   Added InvTWHalf (2/Table1.width) and InvTHHalf (2/Table1.height) computed once in Table1_Init.
   Pan/AudioFade now multiply by inverse instead of dividing per call.
   At 60Hz with 7 balls: ~840 divisions/sec eliminated.

4. BallVel ^2 ELIMINATION (line ~909)
   Replaced ball.VelX^2 + ball.VelY^2 with vx*vx + vy*vy.
   Caches VelX/VelY into locals.
   At 60Hz with 7 balls: ~420 ^2 ops/sec eliminated.

5. Vol ^2 ELIMINATION (line ~868)
   Replaced BallVel(ball)^2 with bv*bv using cached result.
   At 60Hz with 7 balls: ~420 ^2 ops/sec eliminated.

6. OnBallBallCollision ^2 ELIMINATION (line ~923)
   Replaced velocity^2 with velocity*velocity.

7. PRE-BUILT ROLLING SOUND STRINGS (line ~71)
   Pre-built BallRollStr(0..6) = "fx_ballrolling0".."fx_ballrolling6" at module init.
   Eliminates string concatenation inside RollingSoundUpdateBOT.
   At 60Hz with 7 balls: ~420 string allocs/sec eliminated.

8. RollingSoundUpdateBOT INLINE REWRITE (line ~1029)
   Inlined BallVel/Vol/Pitch computations directly in the loop body.
   Cached BOT(b).VelX/VelY/z into locals.
   Cached UBound(BOT) into local ubBot.
   At 60Hz with 7 balls: ~1,260 redundant COM reads/sec + ~420 function call overheads/sec.

9. BallShadowUpdateBOT COM CACHING (line ~1078)
   Cached BOT(b).X, .Y, .Z into locals bx, by, bz.
   Pre-computed BallsizeD6 = Ballsize/6 at module level.
   Added .visible write guards (only writes when value changes).
   Cached UBound(BOT) into local.
   At 60Hz with 7 balls: ~1,260 COM reads/sec + ~420 redundant writes/sec eliminated.

10. FadeLamp / FadeModLamp TypeName() ELIMINATION (lines ~769, ~785)
    Replaced TypeName(object) = "Light"/"Flasher" checks with error trapping
    (attempt .state access; if error, treat as Flasher and use .visible).
    TypeName does COM reflection + string allocation.
    FadeLamp called ~8 times per LampTimer_Timer, FadeModLamp called ~32 times.
    LampTimer runs at 100Hz (10ms interval).
    At 100Hz: ~8,000 TypeName string allocs/sec eliminated.

11. MechsUpdate FLIPPER DELTA GUARDS (line ~1125)
    Cached LeftFlipper.CurrentAngle, RightFlipper.CurrentAngle into locals.
    Added lastLFAngle/lastRFAngle tracking variables.
    Only writes RotZ when angle has actually changed.
    When flippers are at rest (majority of gameplay): eliminates 2 COM writes/tick.
    At 60Hz: ~120 COM writes/sec eliminated when flippers stable.

12. UpdateGI stepScale PRE-COMPUTATION (line ~819)
    Pre-computed stepScale = 0.125 * step once at the top of UpdateGI,
    reused across all For-Each loops that previously computed step/8 inline.

13. EnterPoA_timer myball.X CACHING (line ~444)
    Cached myball.X into local mbx. Previously read twice per tick
    (once for BallShadow.X calculation, once for pan offset).
    Also uses pre-computed BallsizeD6 instead of BallSize/6 per tick.


ESTIMATED TOTAL SAVINGS
========================
- GetBalls: 60 COM calls/sec eliminated
- Exponentiation: ~1,680 ^10 and ^2 ops/sec eliminated
- Division: ~840 divisions/sec eliminated
- COM reads: ~2,520 COM reads/sec eliminated
- COM writes: ~540 redundant writes/sec eliminated
- String allocs: ~8,420 string allocs/sec eliminated (mostly TypeName)
- Function calls: ~420 function call overheads/sec eliminated

Conservative total: ~14,480 redundant operations/sec eliminated.
