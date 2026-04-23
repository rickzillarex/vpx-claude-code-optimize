JP's The Walking Dead VBS — Optimization Changes
===============================================

1. FLIPPER TIMER 1ms to 10ms (line ~613)
   Changed LeftFlipper.TimerInterval from 1 to 10. Flipper tricks timer
   was firing at 1000Hz; 100Hz is visually identical for EOS/SOS/LiveCatch.
   At 1000Hz: eliminates ~900 unnecessary timer calls/sec.

2. FLIPPER TIMER COM CACHING (line ~616)
   Cached LeftFlipper.CurrentAngle, .StartAngle into locals (LFca, LFsa) and
   RightFlipper.CurrentAngle, .StartAngle into (RFca, RFsa) at top of
   LeftFlipper_Timer. Previously each property was read 2-3x per call.
   At 100Hz: ~400 COM reads/sec eliminated.

3. FLIPPER ANIMATE DELTA GUARDS (line ~148)
   Added lastLFAngle/lastRFAngle tracking variables. RotZ write only occurs
   when CurrentAngle actually changes. Flippers at rest skip all COM writes.
   At 60fps: ~120 COM writes/sec eliminated when flippers idle.

4. EXPONENTIATION ELIMINATION — Pan/AudioFade ^10 (line ~971)
   Replaced tmp^10 with chained multiply: t2=tmp*tmp, t4=t2*t2, t8=t4*t4,
   result=t8*t2. Eliminates Exp(10*Log(x)) per call.
   At ~200 calls/sec: ~400 ^10 ops/sec eliminated.

5. EXPONENTIATION ELIMINATION — BallVel ^2 (line ~989)
   Replaced ball.VelX^2 + ball.VelY^2 with vx*vx + vy*vy. Caches VelX/VelY
   into locals, eliminating 2 COM reads per call.
   At ~200 calls/sec: ~400 ^2 ops + ~400 COM reads/sec eliminated.

6. EXPONENTIATION ELIMINATION — Vol ^2 (line ~965)
   Changed BallVel(ball)^2 to bv*bv with cached local.
   At ~100 calls/sec: ~100 ^2 ops/sec eliminated.

7. PRE-COMPUTED INVERSE MULTIPLIERS (line ~962)
   Added InvTWHalf = 2/TableWidth and InvTHHalf = 2/TableHeight at module
   level. Pan and AudioFade now multiply instead of divide.
   At ~200 calls/sec: ~400 divisions/sec eliminated.

8. PRE-BUILT BallRollStr ARRAY (line ~1021)
   Pre-built "fx_ballrolling0" through "fx_ballrolling19" at init.
   Eliminates "fx_ballrolling" & b string concatenation in RollingUpdate.
   At 60fps x 4 balls: ~480 string allocs/sec eliminated.

9. PRE-COMPUTED BS_d2 CONSTANT (line ~1024)
   BallSize/2 computed once at init for shadow height calc.
   At ~240 calls/sec: ~240 divisions/sec eliminated.

10. RollingUpdate COM CACHING + INLINE (line ~1040)
    Cached BOT(b).X/Y/Z/VelX/VelY/VelZ into locals (bx,by,bz,bvx,bvy,bvz).
    Inlined BallVel, Vol, Pitch calculations using cached locals instead of
    calling helper functions that re-read COM.
    Cached UBound(BOT) once at sub entry.
    At 60fps x 4 balls: ~2,400 COM reads/sec + ~480 Sqr/sec eliminated.

11. RollingUpdate DROP SOUND FIX (line ~1075)
    Split VBScript And chain into nested If (VBS And does NOT short-circuit).
    Used cached bvz/bz instead of re-reading COM.
    At 60fps x 4 balls: ~480 COM reads/sec eliminated.

12. LampTimer_Timer 2D ARRAY CACHING (line ~674)
    Cached chgLamp(ii,0) and chgLamp(ii,1) into cIdx/cVal locals at loop top.
    Eliminates double 2D array dereference per lamp change.
    At 50Hz x 20 lamps: ~2,000 2D derefs/sec eliminated.

13. OnBallBallCollision ^2 ELIMINATION (line ~1117)
    Replaced velocity^2 with velocity*velocity.

ESTIMATED TOTAL SAVINGS
========================
Conservative total: ~8,000 redundant operations/sec eliminated.
Peak (multiball + lamp storms): ~12,000+ ops/sec eliminated.
Biggest single win: flipper timer 1ms to 10ms (~900 calls/sec eliminated).
