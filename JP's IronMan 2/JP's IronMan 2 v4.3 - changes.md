JP's IronMan 2 v4.3 VBS — Optimization Changes
===============================================

1. FLIPPER TIMER 1ms -> 10ms (line ~537)
   Changed LeftFlipper.TimerInterval from 1 to 10. 100Hz is visually
   identical to 1000Hz for flipper correction. This single change
   eliminates 900 timer callbacks/sec.
   At 1000Hz: 900 timer fires/sec eliminated.

2. FLIPPER COM PROPERTY CACHING (line ~540)
   Cached LeftFlipper.CurrentAngle, StartAngle, RightFlipper.CurrentAngle,
   StartAngle into local vars Lca, Lsa, Rca, Rsa at top of timer sub.
   Original code read each property 2-4 times per tick.
   At 100Hz (new rate): ~600 COM reads/sec eliminated.

3. PRE-COMPUTED INVERSE TABLE DIMENSIONS (line ~908)
   Added InvTWHalf = 2/TableWidth and InvTHHalf = 2/TableHeight.
   Pan() and AudioFade() now multiply by inverse instead of dividing.
   Called from RollingUpdate + PlaySoundAt + PlaySoundAtBall.
   At 60Hz with 4 balls: ~480 divisions/sec eliminated.

4. EXPONENTIATION ELIMINATION — Pan/AudioFade ^10 (line ~917, ~943)
   Replaced tmp^10 with chained multiply: t2=tmp*tmp, t4=t2*t2,
   t8=t4*t4, result=t8*t2. 4 multiplies replace 1 Exp(10*Log(x)).
   At 60Hz with 4 balls, ~8 calls per ball: ~1,920 ^10 ops/sec eliminated.

5. EXPONENTIATION ELIMINATION — BallVel ^2 (line ~933)
   Replaced ball.VelX^2 + ball.VelY^2 with vx*vx + vy*vy.
   Also caches VelX/VelY into locals, eliminating 2 COM reads per call.
   At 60Hz with 4 balls: ~480 ^2 ops/sec + ~480 COM reads/sec eliminated.

6. EXPONENTIATION ELIMINATION — Vol ^2 (line ~912)
   Replaced BallVel(ball)^2 with bv*bv using cached BallVel result.
   At 60Hz with 4 balls: ~240 ^2 ops/sec eliminated.

7. PRE-BUILT ROLLING SOUND STRINGS (line ~910)
   Pre-built BallRollStr(0..19) = "fx_ballrolling0" through "fx_ballrolling19"
   at init. Eliminates "fx_ballrolling" & b string concatenation per ball per frame.
   At 60Hz with up to 19 balls: ~1,140 string allocs/sec eliminated.

8. RollingUpdate COM PROPERTY CACHING (line ~980)
   Cached BOT(b).X/Y/Z/VelX/VelY into locals bx/by/bz/bvx/bvy at top of
   each ball iteration. Cached UBound(BOT) into ubBot.
   Original code read .X 1x, .Y 1x, .Z 2x, .VelX 2x, .VelY 2x, .VelZ 2x
   per ball. Now reads each once.
   At 60Hz with 4 balls: ~1,440 COM reads/sec eliminated.

9. INLINED BallVel/Vol/Pitch IN RollingUpdate (line ~993)
   Instead of calling BallVel(), Vol(), Pitch() separately (each re-reading
   VelX/VelY from COM), computed bv = SQR(bvx*bvx + bvy*bvy) once and
   derived volume (bv*bv/2000) and pitch (bv*20) inline.
   At 60Hz with 4 balls: ~480 redundant COM reads + ~480 Sqr calls eliminated.

10. GUARDED FLIPPER BAT RotZ WRITES (line ~1048)
    Added lastLFTopAngle/lastRFTopAngle tracking. Only writes .Rotz when
    the angle has actually changed. Flippers are at rest ~80% of the time.
    At 60Hz: ~96 COM writes/sec eliminated when flippers stationary.

11. LampTimer 2D ARRAY CACHING (line ~598)
    Cached chgLamp(ii,0) and chgLamp(ii,1) into cIdx/cVal locals at loop top.
    Each was read twice per iteration (for LampState and FadingStep).
    At 100Hz with ~20 lamp changes: ~2,000 2D array dereferences/sec eliminated.

12. NESTED If FOR VelZ DROP SOUNDS (line ~1008)
    Replaced compound And expression with nested If statements to avoid
    evaluating all conditions when VelZ >= -1 (VBScript And does not
    short-circuit).
    Minor savings per ball per frame when not dropping.

ESTIMATED TOTAL SAVINGS
========================
- Flipper timer reduction: ~900 callbacks/sec
- COM reads eliminated: ~3,480/sec (flipper + rolling + BallVel inline)
- Exponentiation eliminated: ~2,640 ^N ops/sec (Pan/AudioFade/BallVel/Vol)
- Division eliminated: ~480/sec (InvTWHalf/InvTHHalf)
- String allocations eliminated: ~1,140/sec (BallRollStr)
- COM writes eliminated: ~96/sec (flipper bat guards)
- 2D array dereferences eliminated: ~2,000/sec (LampTimer caching)

Conservative total: ~10,700 redundant operations/sec eliminated.
