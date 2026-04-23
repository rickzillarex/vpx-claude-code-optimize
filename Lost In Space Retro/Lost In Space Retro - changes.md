Lost In Space Retro VBS — Optimization Changes
===============================================

1. PRE-COMPUTED INVERSE MULTIPLIERS (line ~190)
   Added InvTblHeight (2/table1.height) and InvTblWidth (2/table1.width)
   at module level, assigned via InitLISConstants() called shortly after
   table init via vpmTimer.AddTimer. Replaces per-call division in
   AudioFade() and AudioPan().
   At 100Hz with 4 balls: ~800 divisions/sec eliminated.

2. AudioFade/AudioPan ^10 ELIMINATION (line ~193)
   Replaced tmp^10 with chained multiply (t2=tmp*tmp, t4=t2*t2,
   t8=t4*t4, result=t8*t2). Handles negative branch separately.
   At 100Hz, called ~8-10x/tick: ~1,800 exponentiation ops/sec eliminated.

3. BallVel ^2 ELIMINATION + COM CACHING (line ~222)
   Replaced ball.VelX^2 + ball.VelY^2 with vx*vx + vy*vy. Caches
   VelX/VelY into locals.
   ~1,000 ^2 ops/sec eliminated.

4. Vol ^2 ELIMINATION (line ~214)
   Replaced BallVel(ball)^2 with bv*bv using cached BallVel result.
   ~400 ^2 ops/sec eliminated.

5. PRE-BUILT BALL ROLL STRINGS (line ~238)
   BallRollStr(0..7) array pre-built at module level. Eliminates
   "fx_ballrolling" & bb string concatenation in RollingTimer_Timer.
   At 100Hz with 4 balls: ~800 string allocs/sec eliminated.

6. RollingTimer_Timer REWRITE (line ~244)
   - Cache UBound(BOT) into ubBot local.
   - Set ball = BOT(bb) to cache ball object reference.
   - Inline BallVel/Vol/Pitch: cache ball.VelX, VelY, z into locals.
     Reduces from ~10 COM reads/ball to 3 COM reads/ball.
   - Use pre-built BallRollStr(bb) for sound names.
   - Removed empty ball drop sound block (had no PlaySound call).
   At 100Hz with 4 balls: ~2,800 COM reads/sec eliminated.

7. FlipperTimer_Timer COM CACHING (line ~285)
   Cached LeftFlipper.currentangle and RightFlipper.currentangle into
   locals caL/caR. Each was read twice (once for shadow, once for logo).
   Now 2 COM reads instead of 4.
   At ~60Hz (frame rate): ~120 COM reads/sec eliminated.

8. OnBallBallCollision ^2 ELIMINATION (line ~277)
   Replaced velocity^2 with velocity*velocity.

ESTIMATED TOTAL SAVINGS
========================
Conservative total: ~7,720 redundant operations/sec eliminated.
Breakdown: ~2,920 COM reads, ~1,800 exponentiations, ~1,400 ^2 ops,
~800 divisions, ~800 string allocs eliminated per second.
