Police Force VBS — Optimization Changes
===============================================

1. PRE-COMPUTED INVERSE MULTIPLIERS (line ~608)
   Added InvPolHeight (2/Police.height) and InvPolWidth (2/Police.width)
   computed once in InitPoliceConstants(), called from Police_init.
   Replaces per-call division in AudioFade() and AudioPan().
   At 100Hz with 2-4 balls: ~600 divisions/sec eliminated.

2. AudioFade/AudioPan ^10 ELIMINATION (line ~615)
   Replaced tmp^10 with chained multiply (t2=tmp*tmp, t4=t2*t2,
   t8=t4*t4, result=t8*t2). Handles negative branch separately.
   At 100Hz, called ~6-8x/tick: ~1,400 exponentiation ops/sec eliminated.

3. BallVel ^2 ELIMINATION + COM CACHING (line ~645)
   Replaced ball.VelX^2 + ball.VelY^2 with vx*vx + vy*vy. Caches
   VelX/VelY into locals (2 COM reads instead of 2, but eliminates
   exponentiation overhead).
   ~800 ^2 ops/sec eliminated.

4. Vol ^2 ELIMINATION (line ~637)
   Replaced BallVel(ball)^2 with bv*bv using cached BallVel result.
   ~300 ^2 ops/sec eliminated.

5. PRE-BUILT STRING ARRAYS (line ~660)
   BallRollStr(0..4) and BallDropStr(0..4) pre-built at module level.
   Eliminates "fx_ballrolling" & b and "fx_ball_drop" & b concatenation.
   At 100Hz with 2 balls: ~400 string allocs/sec eliminated.

6. RollingTimer_Timer REWRITE (line ~670)
   - Cache UBound(BOT) into ubBot local.
   - Set ball = BOT(b) to cache ball object reference.
   - Inline BallVel: cache ball.VelX, VelY, VelZ, z into locals.
     Reduces from ~12 COM reads/ball to 4 COM reads/ball.
   - Use pre-built BallRollStr(b) and BallDropStr(b).
   - Ball drop sound uses cached bvz/bz instead of re-reading COM.
   At 100Hz with 2 balls: ~1,600 COM reads/sec eliminated.

7. DisplayTimer_Timer VISIBLE GUARDS (line ~329)
   Added module-level last-state tracking (lastL23St, lastZ42St, etc.)
   for 6 lamp-to-primitive visible mirrors (Jackpot, BMan, GMan, Oman,
   YMan, Wman). Only writes .visible when source lamp state changes.
   Previously wrote 12 unconditional visible assignments per tick
   (6 if/then pairs always executing one branch).
   At ~60Hz: ~720 redundant COM writes/sec eliminated when lamps stable.

8. OnBallBallCollision ^2 ELIMINATION (line ~707)
   Replaced velocity^2 with velocity*velocity.

ESTIMATED TOTAL SAVINGS
========================
Conservative total: ~5,820 redundant operations/sec eliminated.
Breakdown: ~1,600 COM reads, ~720 COM writes, ~1,400 exponentiations,
~1,100 ^2 ops, ~600 divisions, ~400 string allocs eliminated per second.
