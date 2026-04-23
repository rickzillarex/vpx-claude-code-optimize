Secret Service VBS -- Optimization Changes
===============================================

1. LAMP TIMER INTERVAL 5ms -> 16ms (line ~404)
   Changed LampTimer.Interval from 5 to 16 (frame-rate aligned).
   Lamp fading is handled by VP's built-in light states (NFadeL just sets state 0/1),
   so 62Hz is visually identical to 200Hz. At 200Hz with ~70 lamp calls: ~14,000 ops/sec.
   At 62Hz: ~4,375 ops/sec.

2. LAMPTIMER chglamp CACHING (line ~409)
   Cached chgLamp(ii,0) and chgLamp(ii,1) into local cIdx/cVal variables at loop top.
   Eliminates redundant 2D array dereferences per lamp change.
   At ~20 lamp changes/batch: ~40 redundant 2D dereferences/tick eliminated.

3. AUDIOFADE/AUDIOPAN/PAN ^10 ELIMINATION (line ~821)
   Replaced tmp^10 with chained multiplies: t2=tmp*tmp, t4=t2*t2, t8=t4*t4, result=t8*t2.
   Applied to AudioFade, AudioPan, and Pan functions.
   Pre-computed InvTWHalf = 2/table1.width and InvTHHalf = 2/table1.height at module level.
   Replaces per-call division with multiplication.
   At 100Hz with 5 balls, each calling AudioFade+AudioPan: ~1,000 ^10 ops/sec eliminated,
   ~1,000 divisions/sec eliminated.

4. VOL ^2 ELIMINATION (line ~849)
   Changed Vol = BallVel(ball)^2/VolDiv to bv*bv/VolDiv. Eliminates Exp/Log ^2 overhead.

5. BALLVEL ^2 ELIMINATION (line ~858)
   Changed ball.VelX^2 + ball.VelY^2 to vx*vx + vy*vy with cached COM reads.
   Eliminates 2 COM reads per call (VelX/VelY cached into locals).

6. BALLROLLING_TIMER FULL REWRITE (line ~872)
   - Pre-built BallRollStr()/BallDropStr() arrays at init. Eliminates "fx_ballrolling" & b
     and "fx_ball_drop" & b string concatenation every tick per ball.
   - Set ball = BOT(b) caches ball object reference. Eliminates repeated array indexing.
   - Cached ball.VelX/VelY/VelZ/z into locals bvx/bvy/bvz/bz. Eliminates ~10 COM reads
     per ball per tick.
   - Inlined BallVel/Vol/Pitch computations directly. Eliminates 3 function calls + 6
     redundant COM reads per ball (BallVel re-reads VelX/VelY each call).
   - Cached UBound(BOT) once.
   - Used nested If instead of And for VelZ/z checks (VBScript no short-circuit).
   At 100Hz with 5 balls: ~3,000 COM reads/sec + ~1,500 string allocs/sec + ~1,500
   function calls/sec eliminated.

7. BALLSHADOWUPDATE COM CACHING + WRITE GUARDS (line ~940)
   - Cached BOT(b).X/Y/Z into locals bx/by/bz. Eliminates ~15 COM reads per ball per tick
     (X/Y/Z were each read 2-3 times).
   - Added .visible write guards: only writes when value changes. Eliminates ~5 redundant
     COM writes per ball per tick when shadows are in stable state.
   - Cached UBound(BOT) once.
   At 60Hz with 5 balls: ~4,500 COM reads/sec + ~300 COM writes/sec eliminated.

8. FLIPPER SHADOW WRITE GUARDS (line ~927)
   Added lastLFAngle/lastRFAngle tracking. Only writes FlipperLSh/FlipperRSh.RotZ when
   angle actually changes. When flippers at rest (~80% of gameplay): eliminates 2 COM
   writes per tick.
   At FlipperTimer frequency: ~120 redundant COM writes/sec eliminated.

9. ONBALLBALLCOLLISION ^2 ELIMINATION (line ~919)
   Changed velocity^2 to velocity*velocity. Minor per-event savings.


ESTIMATED TOTAL SAVINGS
========================
Conservative total: ~12,000 redundant operations/sec eliminated.
Primary wins: BallRolling inlining (~5,000), BallShadow COM caching (~4,800),
AudioFade/Pan ^10 elimination (~2,000), LampTimer interval reduction (~9,600 skipped
iterations/sec).
