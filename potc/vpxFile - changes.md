POTC VBS — Optimization Changes
===============================================

1. PRE-COMPUTED CONSTANTS (line ~23)
   Added InitPrecomputedConstants() called from Table1_Init. Pre-computes
   InvTblHeight (2/table1.height), InvTblWidth (2/table1.width),
   HalfTblWidth (Table1.Width/2), BallSizeDiv6 (Ballsize/6),
   InvFlashScale (1/(2.55*100)). Eliminates repeated division in
   AudioFade, AudioPan, BallShadowUpdate, and Flash subs.
   At 100Hz across all callers: ~2,000 divisions/sec eliminated.

2. PRE-BUILT BALL ROLL STRINGS (line ~23)
   BallRollStr(0..5) array pre-built at init with "fx_ballrolling" & i.
   Eliminates string concatenation in RollingTimer_Timer per ball per tick.
   At 100Hz with 5 balls: ~1,000 string allocs/sec eliminated.

3. AudioFade/AudioPan ^10 ELIMINATION (line ~700)
   Replaced tmp^10 with chained multiply: t2=tmp*tmp, t4=t2*t2, t8=t4*t4,
   result=t8*t2. Handles negative branch separately.
   At 100Hz, called ~10x/tick (rolling + shadow + collisions):
   ~2,000 exponentiation ops/sec eliminated.

4. BallVel ^2 ELIMINATION (line ~735)
   Replaced ball.VelX^2 + ball.VelY^2 with vx*vx + vy*vy. Caches
   VelX/VelY into locals (2 COM reads instead of 2).
   Called from Vol, Pitch, RollingTimer: ~1,500 ^2 ops/sec eliminated.

5. RollingTimer_Timer REWRITE (line ~753)
   - Cache UBound(BOT) into ubBot local (1 call instead of 3).
   - Use pre-built BallRollStr(b) instead of "fx_ballrolling" & b.
   - Set ball = BOT(b) to cache ball object reference.
   - Inline BallVel/Vol/Pitch: cache ball.VelX, VelY, z into locals.
     Reduces from ~10 COM reads/ball to 3 COM reads/ball.
   At 100Hz with 4 balls: ~2,800 COM reads/sec eliminated.

6. BallShadowUpdate REWRITE (line ~800)
   - Cache UBound(BOT) into local.
   - Cache BOT(b).X/Y/Z into bx/by/bz locals (3 reads instead of ~10).
   - Use pre-computed HalfTblWidth and BallSizeDiv6.
   - Guard .visible writes: only write when value changes.
   At 100Hz with 4 balls: ~2,800 COM reads/sec + ~400 redundant
   .visible writes/sec eliminated.

7. MiscTimer_Timer VISIBLE GUARDS (line ~593)
   Guard all 8 lamp-mirror .visible writes (GISpot1b/2b/3b, L24b, L32b,
   L40b, L48b, L56b) with last-state tracking. Only write COM property
   when the source lamp state has actually changed.
   At 100Hz: ~800 redundant COM writes/sec eliminated when lamps stable.

8. FLASH SUB OPTIMIZATION (line ~530)
   Replaced (level/2.55)/100 with level * InvFlashScale (pre-computed
   inverse). Flash22 and Flash30 compute scale once into local, reuse
   across 3-5 objects instead of recomputing per object.
   Flash22: 2 divisions eliminated per call. Flash30: 4 divisions eliminated.
   At ~50 calls/sec: ~200 divisions/sec eliminated.

9. OnBallBallCollision ^2 ELIMINATION (line ~785)
   Replaced velocity^2 with velocity*velocity.

ESTIMATED TOTAL SAVINGS
========================
Conservative total: ~11,300 redundant operations/sec eliminated.
Breakdown: ~2,200 divisions, ~2,000 exponentiations, ~5,600 COM reads,
~1,200 COM writes, ~1,000 string allocations avoided per second.
