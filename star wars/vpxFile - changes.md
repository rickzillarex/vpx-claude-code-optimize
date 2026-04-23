Star Wars VBS — Optimization Changes
===============================================

1. PRE-COMPUTED TRIG CONSTANTS (line ~431)
   Added CPI and PIover180 at module level. dSin/dCos now use PIover180
   instead of calling Pi() function (which computes 4*Atn(1) every call)
   and dividing by 180 each time.
   TiR2Head_Timer calls dSin 2x + dCos 2x per tick.
   At ~100Hz: ~400 Pi() function calls + ~400 divisions/sec eliminated.

2. EXPONENTIATION ELIMINATION — AudioFade/AudioPan ^10 (line ~837)
   Replaced tmp^10 with chained multiply: t2=tmp*tmp, t4=t2*t2, t8=t4*t4,
   result=t8*t2. Eliminates Exp(10*Log(x)) per call.
   At ~200 calls/sec (rolling + shadows + collide + hit sounds):
   ~400 ^10 ops/sec eliminated.

3. EXPONENTIATION ELIMINATION — BallVel ^2 (line ~866)
   Replaced ball.VelX^2 + ball.VelY^2 with vx*vx + vy*vy. Caches VelX/VelY
   into locals, eliminating 2 COM reads per call.
   At ~200 calls/sec: ~400 ^2 ops + ~400 COM reads/sec eliminated.

4. EXPONENTIATION ELIMINATION — Vol ^2 (line ~857)
   Changed BallVel(ball)^2 to bv*bv with cached local.
   At ~100 calls/sec: ~100 ^2 ops/sec eliminated.

5. PRE-COMPUTED INVERSE MULTIPLIERS (line ~834)
   Added InvTWHalf = 2/table1.width and InvTHHalf = 2/table1.height at
   module level. AudioPan and AudioFade now multiply instead of divide.
   At ~200 calls/sec: ~400 divisions/sec eliminated.

6. PRE-BUILT BallRollStr + BallDropStr ARRAYS (line ~871)
   Pre-built "fx_ballrolling0" through "fx_ballrolling6" and
   "fx_ball_drop0" through "fx_ball_drop6" at init.
   Eliminates string concatenation in RollingTimer_Timer.
   At 60fps x 3 balls: ~360 string allocs/sec eliminated.

7. RollingTimer_Timer COM CACHING + INLINE (line ~890)
   Cached BOT(b).VelX/VelY/VelZ/z into locals (bvx,bvy,bvz,bz).
   Inlined BallVel, Vol, Pitch calculations using cached locals instead
   of calling helper functions that re-read COM + re-compute Sqr.
   Cached UBound(BOT) once at sub entry.
   At 60fps x 3 balls: ~1,800 COM reads/sec + ~360 Sqr/sec eliminated.

8. RollingTimer_Timer DROP SOUND FIX (line ~920)
   Split VBScript And chain into nested If (VBS And does NOT short-circuit).
   Used cached bvz/bz instead of re-reading COM.
   At 60fps x 3 balls: ~360 COM reads/sec eliminated.

9. FLIPPER SHADOW DELTA GUARDS — RealTime_Timer (line ~933)
   Added lastLFSAngle/lastRFSAngle tracking. RotZ write on flipper shadow
   only occurs when CurrentAngle actually changes. Flippers at rest skip
   all COM writes.
   At 60fps: ~120 COM writes/sec eliminated when flippers idle.

10. BallShadowUpdate COM CACHING (line ~946)
    Cached BOT(b).X/Y/Z into locals (bx,by,bz) at top of each iteration.
    Previously .X/.Y/.Z read 3-5x per ball per frame.
    Cached UBound(BOT) once at sub entry.
    At 60fps x 3 balls: ~900 COM reads/sec eliminated.

11. BallShadowUpdate .visible GUARD (line ~956)
    Added delta guard: only writes .visible when value actually changes.
    During normal play, balls are almost always visible (z between 20-200).
    At 60fps x 3 balls: ~180 redundant COM writes/sec eliminated.

12. BallShadow ARRAY PRE-BUILT (line ~930)
    Moved BallShadow = Array(...) to a one-time init function instead of
    recreating the Array() every single frame in BallShadowUpdate.
    At 60fps: ~60 Array allocations/sec eliminated.

13. LampTimer_Timer 2D ARRAY CACHING (line ~527)
    Cached chgLamp(ii,0) and chgLamp(ii,1) into cIdx/cVal locals at loop top.
    Eliminates triple 2D array dereference per lamp change (was read 3x each).
    At 29Hz x 20 lamps: ~1,740 2D derefs/sec eliminated.

14. OnBallBallCollision ^2 ELIMINATION (line ~923)
    Replaced velocity^2 with velocity*velocity.

ESTIMATED TOTAL SAVINGS
========================
Conservative total: ~7,500 redundant operations/sec eliminated.
Peak (multiball + lamp storms): ~10,000+ ops/sec eliminated.
Biggest single wins: Pi() function elimination (~400 calls/sec), ^10 chain
multiply (~400 ops/sec), RollingTimer COM caching (~1,800 reads/sec).
