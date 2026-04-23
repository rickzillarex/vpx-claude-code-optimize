JP's World Joker Tour v600 VBS — Optimization Changes
===============================================

1. FLIPPER TIMER 1ms -> 10ms (line ~619)
   Changed LeftFlipper.TimerInterval from 1 to 10. 100Hz is visually
   identical to 1000Hz for flipper tricks. This single change eliminates
   900 timer calls/sec and all downstream COM reads within them.
   At 1000Hz: ~4,500 function calls/sec eliminated.

2. FLIPPER TIMER COM CACHING (line ~622)
   Cached LeftFlipper.CurrentAngle, .StartAngle, .EndAngle and
   RightFlipper.CurrentAngle, .StartAngle, .EndAngle into locals
   at top of LeftFlipper_Timer. These properties were read 2-4 times
   each per call. Now 3 COM reads per flipper instead of 6-8.
   At 100Hz (after interval fix): ~600 COM reads/sec eliminated.
   (At original 1000Hz this would have been ~6,000/sec.)

3. PRE-COMPUTED INVERSE MULTIPLIERS (line ~691)
   Added InvTWHalf (2/TableWidth) and InvTHHalf (2/TableHeight) computed
   once at module level. Replaces division per call in Pan() and AudioFade().
   At 100Hz with 4 balls: ~800 divisions/sec eliminated.

4. PRE-BUILT BALL ROLL STRINGS (line ~693)
   BallRollStr(0..19) array pre-built at module level. Eliminates
   "fx_ballrolling" & b string concatenation in RollingTimer_Timer.
   At 100Hz with 4+ balls: ~800 string allocs/sec eliminated.

5. AudioFade/Pan ^10 ELIMINATION (line ~700)
   Replaced tmp^10 with chained multiply (t2*t2=t4, t4*t4=t8, t8*t2=t10).
   Handles negative branch separately. Applied to both Pan() and AudioFade().
   At 100Hz, called ~8-12x/tick: ~2,000 exponentiation ops/sec eliminated.

6. BallVel ^2 ELIMINATION + COM CACHING (line ~714)
   Replaced ball.VelX^2 + ball.VelY^2 with vx*vx + vy*vy. Caches
   VelX/VelY into locals.
   Called from Vol, Pitch, RollingTimer: ~1,200 ^2 ops/sec eliminated.

7. Vol ^2 ELIMINATION (line ~696)
   Replaced BallVel(ball)^2 with bv*bv using cached BallVel result.
   ~400 ^2 ops/sec eliminated.

8. debug.print REMOVAL (line ~235)
   Removed debug.print from GIUpdate callback. debug.print allocates COM
   strings and writes to debug output on every GI change event.
   Eliminates string allocation on every GI callback.

9. RollingTimer_Timer REWRITE (line ~763)
   - Cache UBound(BOT) into ubBot local.
   - Set ball = BOT(b) to cache ball object reference.
   - Inline BallVel: cache ball.VelX, VelY, VelZ, z into locals.
     Reduces from ~15 COM reads/ball to 4 COM reads/ball.
   - Use pre-built BallRollStr(b) for sound names.
   - Dropping sounds use cached bvz/bz instead of re-reading COM.
   - Speed control uses cached bvx/bvy instead of re-reading COM.
   At 100Hz with 4 balls: ~4,400 COM reads/sec eliminated.

10. OnBallBallCollision ^2 ELIMINATION (line ~822)
    Replaced velocity^2 with velocity*velocity.

ESTIMATED TOTAL SAVINGS
========================
Conservative total: ~14,700 redundant operations/sec eliminated.
Breakdown: ~4,500 timer calls (from 1ms->10ms), ~5,000 COM reads,
~2,000 exponentiations, ~1,200 ^2 ops, ~800 divisions,
~800 string allocs, ~400 misc ops eliminated per second.
