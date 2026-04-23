Ripley's Believe It Or Not VBS — Optimization Changes
===============================================

1. FLIPPER TIMER 1ms to 10ms (line ~331)
   Changed LeftFlipper.TimerInterval from 1 to 10. At 1ms the flipper tricks
   timer fired 1000Hz; 10ms (100Hz) is visually identical for EOS/SOS logic.
   At 1000Hz: eliminates ~900 unnecessary timer calls/sec.

2. FLIPPER TIMER COM CACHING (line ~334)
   Cached LeftFlipper.CurrentAngle, .StartAngle into locals (LFca, LFsa) and
   RightFlipper.CurrentAngle, .StartAngle into (RFca, RFsa) at top of
   LeftFlipper_Timer. Previously each property was read 2-3x per call.
   At 100Hz: ~400 COM reads/sec eliminated.

3. FLIPPER ANIMATE DELTA GUARDS (line ~63)
   Added lastLFAngle/lastRFAngle/lastRF1Angle tracking. RotZ write only
   occurs when CurrentAngle actually changes. Flippers at rest (80%+ of time)
   skip all COM writes.
   At 60fps: ~180 COM writes/sec eliminated when flippers idle.

4. EXPONENTIATION ELIMINATION — Pan/AudioFade ^10 (line ~957)
   Replaced tmp^10 with chained multiply: t2=tmp*tmp, t4=t2*t2, t8=t4*t4,
   result=t8*t2. Eliminates Exp(10*Log(x)) per call.
   At ~200 calls/sec (rolling + drop + collide): ~400 ^10 ops/sec eliminated.

5. EXPONENTIATION ELIMINATION — BallVel ^2 (line ~975)
   Replaced ball.VelX^2 + ball.VelY^2 with vx*vx + vy*vy. Also caches
   VelX/VelY into locals, eliminating 2 COM reads per call.
   At ~200 calls/sec: ~400 ^2 ops + ~400 COM reads/sec eliminated.

6. EXPONENTIATION ELIMINATION — Vol ^2 (line ~950)
   Changed BallVel(ball)^2 to bv*bv with cached local.
   At ~100 calls/sec: ~100 ^2 ops/sec eliminated.

7. PRE-COMPUTED INVERSE MULTIPLIERS (line ~943)
   Added InvTWHalf = 2/TableWidth and InvTHHalf = 2/TableHeight at module
   level. Pan and AudioFade now multiply instead of divide.
   At ~200 calls/sec: ~400 divisions/sec eliminated.

8. PRE-COMPUTED BS_d2 CONSTANT (line ~946)
   BallSize/2 computed once at init. Used in RollingUpdate shadow height calc.
   At ~100 calls/sec: ~100 divisions/sec eliminated.

9. PRE-BUILT BallRollStr ARRAY (line ~1007)
   Pre-built "fx_ballrolling0" through "fx_ballrolling19" at init.
   Eliminates "fx_ballrolling" & b string concatenation in RollingUpdate.
   At 60fps x 4 balls: ~480 string allocs/sec eliminated.

10. RollingUpdate COM CACHING + INLINE (line ~1020)
    Cached BOT(b).X/Y/Z/VelX/VelY/VelZ into locals (bx,by,bz,bvx,bvy,bvz).
    Inlined BallVel, Vol, Pitch calculations using cached locals.
    Previously: BallVel called 3x per ball (velocity check, Pitch, Vol), each
    re-reading VelX/VelY from COM + computing Sqr.
    Cached UBound(BOT) once at sub entry.
    At 60fps x 4 balls: ~2,400 COM reads/sec + ~480 Sqr/sec eliminated.

11. RollingUpdate DROP SOUND FIX (line ~1053)
    Split VBScript And chain into nested If (VBS And does NOT short-circuit).
    Used cached bvz/bz instead of re-reading COM.
    At 60fps x 4 balls: ~480 COM reads/sec eliminated.

12. LampTimer_Timer 2D ARRAY CACHING (line ~650)
    Cached chgLamp(ii,0) and chgLamp(ii,1) into cIdx/cVal locals at loop top.
    Eliminates double 2D array dereference per lamp change.
    At 50Hz x 20 lamps: ~2,000 2D derefs/sec eliminated.

13. OnBallBallCollision ^2 ELIMINATION (line ~1085)
    Replaced velocity^2 with velocity*velocity.
    Low frequency but free to fix.

ESTIMATED TOTAL SAVINGS
========================
Conservative total: ~7,500 redundant operations/sec eliminated.
Peak (multiball + lamp storms): ~12,000+ ops/sec eliminated.
Biggest single win: flipper timer 1ms to 10ms (~900 calls/sec eliminated).
