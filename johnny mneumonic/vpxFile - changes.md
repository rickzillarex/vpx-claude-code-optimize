johnnyvpxFile VBS — Optimization Changes
===============================================

1. PRE-COMPUTED TABLE DIMENSION INVERSES (line ~994)
   Cached JM.width and JM.height into JMWidth/JMHeight at module level.
   Pre-computed InvJMWHalf = 2/JMWidth and InvJMHHalf = 2/JMHeight.
   AudioFade/AudioPan now multiply by inverse instead of dividing by
   JM.width/JM.height (which are COM reads + division each call).
   At 60Hz with 4 balls: ~480 divisions + ~480 COM reads/sec eliminated.

2. EXPONENTIATION ELIMINATION — AudioFade/AudioPan ^10 (line ~1004, ~1018)
   Replaced tmp^10 with chained multiply: t2=tmp*tmp, t4=t2*t2,
   t8=t4*t4, result=t8*t2. 4 multiplies replace 1 Exp(10*Log(x)).
   At 60Hz with 4 balls, ~8 calls per ball: ~1,920 ^10 ops/sec eliminated.

3. EXPONENTIATION ELIMINATION — BallVel ^2 (line ~1040)
   Replaced ball.VelX^2 + ball.VelY^2 with vx*vx + vy*vy.
   Also caches VelX/VelY into locals.
   At 60Hz with 4 balls: ~480 ^2 ops + ~480 COM reads/sec eliminated.

4. EXPONENTIATION ELIMINATION — Vol ^2 (line ~1028)
   Replaced BallVel(ball)^2 with bv*bv using cached result.
   At 60Hz with 4 balls: ~240 ^2 ops/sec eliminated.

5. PRE-BUILT ROLLING SOUND STRINGS (line ~1001)
   Pre-built BallRollStr(0..5) and BallDropStr(0..5) at init.
   Eliminates "fx_ballrolling" & b and "fx_ball_drop" & b string
   concatenation per ball per frame.
   At 60Hz with 4 balls: ~480 string allocs/sec eliminated.

6. TypeName() REDUCTION IN AddLamp (line ~905)
   Original: 3 TypeName() calls per AddLamp invocation (one each for
   Light/Flasher/Primitive). Changed to single TypeName() per call using
   ElseIf chain, and guarded by VarType() == 9 check first.
   22 AddLamp calls per LampTimer tick at 100Hz = 2,200 ticks/sec.
   Reduces from 6,600 TypeName calls/sec to 2,200.
   At 100Hz: ~4,400 TypeName ops/sec eliminated.

7. RollingSoundUpdate COM CACHING + INLINED HELPERS (line ~1055)
   Cached BOT(b).VelX/VelY/z into locals bvx/bvy/bz. Cached UBound(BOT).
   Inlined BallVel, Vol, Pitch computations: computed bv = SQR(bvx*bvx +
   bvy*bvy) once, derived volume (bv*bv/6000) and pitch (bv*20) inline.
   At 60Hz with 4 balls: ~1,440 COM reads + ~480 Sqr calls eliminated.

8. GUARDED FLIPPER BAT/SHADOW RotZ WRITES (line ~942)
   Cached LeftFlipper.CurrentAngle and RightFlipper.CurrentAngle into
   locals, shared between FlipperL/lfs and FlipperR/rfs writes.
   Added lastLFAngle/lastRFAngle delta guards. When flippers are at rest
   (~80% of gameplay), skips 4 COM writes per frame.
   At 60Hz: ~192 COM writes/sec eliminated when flippers stationary.
   Also eliminates duplicate COM reads (was reading .CurrentAngle 2x per
   flipper per frame, now reads once).
   At 60Hz: ~240 COM reads/sec eliminated.

9. NESTED If FOR DROP SOUNDS (line ~1082)
   Replaced compound And expression (VelZ < -1 and z < 55 and z > 27)
   with nested If statements to avoid evaluating all conditions when
   VelZ >= -1 (VBScript And does not short-circuit).

10. OnBallBallCollision ^2 ELIMINATION (line ~1097)
    Replaced velocity^2 with velocity*velocity.

ESTIMATED TOTAL SAVINGS
========================
- TypeName elimination: ~4,400/sec
- Exponentiation eliminated: ~2,640 ^N ops/sec
- COM reads eliminated: ~2,640/sec (table dims + BallVel + flipper + rolling)
- Division eliminated: ~480/sec (InvJMWHalf/InvJMHHalf)
- String allocations eliminated: ~480/sec (BallRollStr + BallDropStr)
- COM writes eliminated: ~192/sec (flipper bat guards)

Conservative total: ~10,800 redundant operations/sec eliminated.
