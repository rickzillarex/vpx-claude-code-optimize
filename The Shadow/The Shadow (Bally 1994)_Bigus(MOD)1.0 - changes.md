The Shadow (Bally 1994) VBS — Optimization Changes
===============================================

1. AUDIOFADE/AUDIOPAN/PAN ^10 → CHAINED MULTIPLY (line ~1449)
   Replaced tmp^10 with t2=tmp*tmp, t4=t2*t2, t8=t4*t4, result=t8*t2
   in AudioFade, AudioPan, and Pan functions. All three rewritten.
   Removed On Error Resume Next (not needed with pre-computed inverse).
   At ~100Hz with 4 balls: ~1,200 exponentiation ops/sec eliminated.

2. AUDIOFADE/AUDIOPAN/PAN DIVISION → PRE-COMPUTED INVERSE (line ~119)
   Replaced `* 2 / table1.height` with `* InvTHHalf` and same for width.
   InvTWHalf/InvTHHalf computed once in Table1_Init.
   At ~100Hz with 4 balls: ~800 divisions/sec eliminated.

3. BallVel ^2 → MULTIPLY + COM CACHING (line ~1493)
   Replaced ball.VelX^2 + ball.VelY^2 with vx*vx + vy*vy. Caches
   VelX/VelY into locals.
   At ~100Hz with 4 balls: ~800 exponent ops + ~800 COM reads/sec eliminated.

4. Vol ^2 → MULTIPLY (line ~1485)
   Replaced BallVel(ball)^2 with bv*bv using cached local.
   At ~100Hz with 4 balls: ~400 exponent ops/sec eliminated.

5. VolZ ^2 → MULTIPLY (line ~1501)
   Replaced BallVelZ(ball)^2 with bvz*bvz using cached local.
   Per call: eliminates 1 exponentiation op.

6. PRE-BUILT STRING ARRAYS (line ~124)
   Pre-built "fx_ballrolling0" through "fx_ballrolling6",
   "fx_metalrolling0" through "fx_metalrolling6", and
   "fx_ball_drop0" through "fx_ball_drop6" at init.
   Eliminates string concatenation in RollingTimer hot loop.
   At ~100Hz with 4 balls: ~2,400 string allocs/sec eliminated.

7. RollingTimer REWRITE — COM CACHING + INLINED HELPERS (line ~1520)
   - Set ball = BOT(b) caches ball object reference
   - Cached ball.VelX/VelY/VelZ/z/x/y into locals
   - Inlined BallVel, Vol, Pitch, Pan, AudioFade using cached locals
   - Eliminated 5+ separate helper calls per ball (each re-reading COM)
   - Used pre-built BallRollStr/MetalRollStr/BallDropStr arrays
   - Cached UBound(BOT) once
   At 100Hz with 4 balls: ~4,000 COM reads/sec + ~2,000 function calls/sec eliminated.

8. BallShadowUpdate COM CACHING (line ~1620)
   - Cached BOT(b).X/Y/Z into bx/by/bz locals (3 reads instead of 7+)
   - Moved BallShadow array creation to module level (built once in Table1_Init)
   - Cached UBound(BOT) once
   At 60Hz with 4 balls: ~960 COM reads/sec eliminated.

9. FlipperTimer GUARDED WRITES (line ~1607)
   - Cache LeftFlipper/RightFlipper/RightFlipper2.currentangle once
   - Added delta guards: skip .RotZ COM writes when angle unchanged
   - 3 flippers (LeftFlipper, RightFlipper, RightFlipper2) all guarded
   At 60Hz: when at rest, ~180 COM writes/sec skipped.

10. UpdateFlipperLogos GUARDED WRITES (line ~282)
    - Cache LeftFlipper/RightFlipper/RightFlipper2.CurrentAngle once
    - Added delta guards: skip .ObjRotZ writes when angle unchanged
    - 3 flipper logos (batleft, batright, batright1) all guarded
    At 60Hz: when at rest, ~180 COM writes/sec skipped.

11. HandleMiniPF FLIPPER/GATE ANGLE CACHING + GUARDED WRITES (3 copies at lines ~1840/1941/2032)
    - Cache LeftFlipper/RightFlipper/RightFlipper2.CurrentAngle into locals
    - Cache Gate1/Gate2.CurrentAngle into locals
    - Added delta guards: skip .RotZ/.RotX writes when angle unchanged
    - Applied to all 3 copies (MiniPlayfieldDifficulty Case 0/1/2)
    Via MotorCallback (~60Hz): 5 COM reads cached per frame.
    When at rest: ~300 COM writes/sec skipped.

12. LampTimer_Timer ARRAY CACHING (line ~1125)
    Cached chgLamp(ii,0) and chgLamp(ii,1) into locals cIdx/cVal.
    Eliminates redundant 2D array dereferences per lamp change.
    At lamp storms: ~100+ 2D dereferences/sec eliminated.

13. OnBallBallCollision ^2 → MULTIPLY (line ~1600)
    Replaced Csng(velocity)^2 with v*v using cached local.
    Per collision: eliminates 1 exponentiation op.


ESTIMATED TOTAL SAVINGS
========================
- COM property reads: ~6,500/sec eliminated
- Exponentiation ops: ~3,200/sec eliminated
- String allocations: ~2,400/sec eliminated
- Division ops: ~800/sec eliminated
- COM writes (guarded): ~660/sec eliminated (when flippers at rest)
- Function calls: ~2,000/sec eliminated
- Array dereferences: ~100/sec eliminated
Conservative total: ~15,600 redundant operations/sec eliminated.
