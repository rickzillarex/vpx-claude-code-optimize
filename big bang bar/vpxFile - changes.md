Big Bang Bar VBS — Optimization Changes
===============================================

1. AUDIOFADE/AUDIOPAN ^10 → CHAINED MULTIPLY (line ~1771)
   Replaced tmp^10 with t2=tmp*tmp, t4=t2*t2, t8=t4*t4, result=t8*t2
   in AudioFade, AudioPan, and Pan functions. All three rewritten.
   At ~100Hz with 4 balls: ~1,200 exponentiation ops/sec eliminated.

2. AUDIOFADE/AUDIOPAN DIVISION → PRE-COMPUTED INVERSE (line ~1771)
   Replaced `* 2 / table.height` with `* InvTHHalf` and same for width.
   InvTWHalf/InvTHHalf computed once in Table_Init.
   At ~100Hz with 4 balls: ~800 divisions/sec eliminated.

3. BallVel ^2 → MULTIPLY + COM CACHING (line ~1809)
   Replaced ball.VelX^2 + ball.VelY^2 with vx*vx + vy*vy. Caches
   VelX/VelY into locals.
   At ~100Hz with 4 balls: ~800 exponent ops + ~800 COM reads/sec eliminated.

4. Vol ^2 → MULTIPLY (line ~1801)
   Replaced BallVel(ball)^2 with bv*bv using cached local.
   At ~100Hz with 4 balls: ~400 exponent ops/sec eliminated.

5. PRE-BUILT BallRollStr ARRAY (line ~1819)
   Pre-built "fx_ballrolling0" through "fx_ballrolling10" at init.
   Eliminates string concatenation in RollingTimer hot loop.
   At ~100Hz with 4 balls: ~800 string allocs/sec eliminated.

6. RollingTimer REWRITE — COM CACHING + INLINED HELPERS (line ~1828)
   - Set ball = BOT(b) caches ball object reference
   - Cached ball.VelX/VelY/z into bvx/bvy/bz locals
   - Inlined BallVel, Vol, Pitch using cached locals
   - Eliminated 3 separate helper calls per ball (each re-reading COM)
   - Cached UBound(BOT) once
   At 100Hz with 4 balls: ~2,400 COM reads/sec + ~1,200 function calls/sec eliminated.

7. BallShadowUpdate COM CACHING (line ~1872)
   - Cached BOT(b).X/Y/Z into bx/by/bz locals (3 reads instead of 7)
   - Moved BallShadow array creation to module level (built once, not per frame)
   - Cached UBound(BOT) once
   At 60Hz with 4 balls: ~960 COM reads/sec eliminated.

8. RealTime_timer FLIPPER ANGLE CACHING + GUARDED WRITES (line ~1580)
   - Cache LeftFlipper.currentangle once (was read 2x for P.roty and lfs.RotZ)
   - Same for RightFlipper (was read 2x)
   - Added delta guards: skip COM writes when angle unchanged
   At 60Hz: ~240 COM reads/sec eliminated. When at rest: ~360 COM writes/sec skipped.

9. LampTimer_Timer ARRAY CACHING (line ~1221)
   Cached chgLamp(ii,0) and chgLamp(ii,1) into locals cIdx/cVal.
   Eliminates redundant 2D array dereferences per lamp change.
   At lamp storms: ~100+ 2D dereferences/sec eliminated.

10. DANCER ANIMATION ARRAY OPTIMIZATION (line ~423)
    Replaced 10-case Select with 100 .visible writes per tick with
    array-indexed frame switching: 2 .visible writes per tick (hide old,
    show new). Pre-built DancerFrames and DancerRotX arrays at module level.
    At dancer timer rate (~60Hz): ~480 redundant COM writes/sec eliminated.

11. OnBallBallCollision ^2 → MULTIPLY (line ~1864)
    Replaced velocity^2 with velocity*velocity. Cached computation
    to avoid duplicate in If/Else branches.
    Per collision: eliminates 1-2 exponentiation ops.


ESTIMATED TOTAL SAVINGS
========================
- COM property reads: ~5,200/sec eliminated
- Exponentiation ops: ~3,200/sec eliminated
- String allocations: ~800/sec eliminated
- Division ops: ~800/sec eliminated
- COM writes (guarded): ~840/sec eliminated
- Function calls: ~1,200/sec eliminated
- Redundant .visible writes: ~480/sec eliminated
Conservative total: ~12,500 redundant operations/sec eliminated.
