AC/DC Pro Vault VBS — Optimization Changes
===============================================

1. SHARED GetBalls IN RealTimeUpdates (line ~1062)
   RealTimeUpdates (MotorCallback, runs every frame at 60Hz) previously called
   RollingSoundUpdate and BallShadowUpdate, each calling GetBalls independently.
   Now calls GetBalls once and passes the BOT array to new BOT-accepting variants
   (RollingSoundUpdateBOT, BallShadowUpdateBOT). Original subs preserved as wrappers.
   At 60Hz: 60 GetBalls COM calls/sec eliminated.

2. EXPONENTIATION ELIMINATION: Pan, AudioFade (lines ~921-938)
   Replaced ^10 with chained multiply (t2=tmp*tmp, t4=t2*t2, t8=t4*t4, result=t8*t2).
   Handles positive/negative branches separately.
   Pan and AudioFade are called from RollingSoundUpdate (per ball per frame),
   PlaySoundAt, PlaySoundAtBall, PlaySoundAtBallVol, etc.
   At 60Hz with 4 balls: ~480 ^10 ops/sec eliminated from rolling alone.

3. PRE-COMPUTED INVERSE TABLE DIMENSIONS (line ~99, init ~141)
   Added InvTWHalf (2/ACDC.width) and InvTHHalf (2/ACDC.height) computed once in ACDC_Init.
   Pan/AudioFade now multiply by inverse instead of dividing per call.
   At 60Hz with 4 balls: ~480 divisions/sec eliminated.

4. BallVel ^2 ELIMINATION (line ~945)
   Replaced ball.VelX^2 + ball.VelY^2 with vx*vx + vy*vy.
   Caches VelX/VelY into locals.
   At 60Hz with 4 balls: ~240 ^2 ops/sec eliminated.

5. Vol ^2 ELIMINATION (line ~917)
   Replaced BallVel(ball)^2 with bv*bv using cached result.
   At 60Hz with 4 balls: ~240 ^2 ops/sec eliminated.

6. OnBallBallCollision ^2 ELIMINATION (line ~1123)
   Replaced velocity^2 with velocity*velocity.

7. PRE-BUILT ROLLING SOUND STRINGS (line ~102)
   Pre-built BallRollStr(0..3) = "fx_ballrolling0".."fx_ballrolling3" at module init.
   Eliminates string concatenation inside RollingSoundUpdate.
   At 60Hz with 4 balls: ~480 string allocs/sec eliminated.

8. RollingSoundUpdate INLINE REWRITE (line ~1084)
   Inlined BallVel/Vol/Pitch computations directly in the loop body.
   Cached BOT(b).VelX/VelY/z into locals.
   Cached UBound(BOT) into local.
   At 60Hz with 4 balls: ~720 redundant COM reads/sec + ~240 function call overheads/sec.

9. BallShadowUpdate COM CACHING (line ~1169)
   Cached BOT(b).X, .Y, .Z into locals bx, by, bz.
   Pre-computed BallsizeD6 = Ballsize/6 at module level.
   Cached ACDC.Width/2 before loop (was computed per ball iteration).
   Added .visible write guards.
   At 60Hz with 4 balls: ~720 COM reads/sec + ~120 redundant writes/sec eliminated.

10. LampTimer_Timer chglamp CACHING (line ~630)
    Cached chglamp(ii,0) and chglamp(ii,1) into local cIdx/cVal at loop top.
    Eliminates repeated 2D array dereference.
    During lamp storms (50-100Hz): ~200-400 redundant 2D dereferences/sec eliminated.

11. LampMod TypeName() ELIMINATION (line ~903)
    Replaced TypeName(object) = "Light"/"Flasher" checks with VarType + error trapping.
    TypeName does COM reflection + string allocation.
    LampMod is called ~40 times per UpdateLamps (every 10ms = 100Hz).
    At 100Hz: ~8,000 TypeName string allocs/sec eliminated.

12. RGBLED TypeName() ELIMINATION (line ~880)
    Same pattern as #11 for the RGBLED function.

13. UpdateCannon_Timer TRIG OPTIMIZATION (line ~366)
    Replaced GPos*Pi/180 with GPos*PIover180 (pre-computed constant).
    Cached -(GPos-90) into local gPosNeg for two LaserR objRotZ assignments.
    At cannon motor speed (~50Hz): ~100 divisions/sec + ~50 redundant computations/sec eliminated.

14. BallInGunRadius ^2 ELIMINATION (line ~329)
    Replaced (Cannon_assyM.X - Sw45.X)^2 with manual multiply.
    One-shot init, minimal savings, but removes last ^2 from module-level code.

15. sw36_timer BELL ANIMATION CACHING (line ~437)
    Cached -SpinnerBell.CurrentAngle into local bellAngle, reused for Bell.RotX,
    Bell_Support.RotX, Bell_Support1.RotX (was reading CurrentAngle 3 times).
    At 100Hz: ~200 COM reads/sec eliminated.

16. PIover180 CONSTANT (line ~97)
    Added PIover180 = Pi/180 as pre-computed constant for trig conversions.


ESTIMATED TOTAL SAVINGS
========================
- GetBalls: 60 COM calls/sec eliminated
- Exponentiation: ~960 ^10 and ^2 ops/sec eliminated
- Division: ~580 divisions/sec eliminated
- COM reads: ~1,840 COM reads/sec eliminated
- COM writes: ~120 redundant writes/sec eliminated
- String allocs: ~8,480 string allocs/sec eliminated (mostly TypeName)
- Function calls: ~240 function call overheads/sec eliminated

Conservative total: ~12,280 redundant operations/sec eliminated.
