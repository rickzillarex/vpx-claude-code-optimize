X-Men vpxFile.vbs — Optimization Changes
===============================================

1. PRE-COMPUTED CONSTANTS (line ~10, assigned in Table_Init ~111)
   Added InvTWHalf (2/Table.width) and InvTHHalf (2/Table.height).
   Declared at module level, assigned inside Table_Init (COM properties not
   available at module-level in VBScript). Eliminates repeated division in
   AudioFade and AudioPan.

2. PRE-BUILT STRING ARRAYS (line ~10)
   Added BallRollStr(6) and BallDropStr(6) arrays.
   Eliminates "fx_ballrolling" & b and "fx_ball_drop" & b string concatenation
   in RollingTimer_Timer hot loop.
   At 100Hz with 6 balls: ~1,800 string allocs/sec eliminated.

3. LAMPTIMER chglamp CACHING (line ~796)
   Cached chgLamp(ii,0) and chgLamp(ii,1) into cIdx/cVal locals at loop top.
   Eliminates 3 redundant 2D array dereferences per lamp per tick (was reading
   chgLamp(ii,0) three times per iteration for LampState, FadingLevel, FlashState).

4. LAMPMOD TypeName OPTIMIZATION (line ~984)
   Changed from 3 separate If+TypeName checks to a single TypeName call cached
   into local `tn`, then ElseIf chain. Reduces from 3 TypeName calls to 1 per
   LampMod invocation. LampMod is called ~30 times per LampTimer tick.
   At 100Hz: ~6,000 TypeName (COM reflection + string alloc) calls/sec eliminated.

5. UPDATEFLIPPERLOGO DELTA GUARDS (line ~1055)
   Added lastLFAngle/lastRFAngle/lastRF1Angle tracking.
   Cached CurrentAngle into locals once per flipper (was reading 2x each: once
   for Logo, once for Shadow). Only writes RotY/RotZ when angle changed.
   At 100Hz: ~600 COM reads/sec + ~600 COM writes/sec eliminated when at rest.

6. AUDIOFADE/AUDIOPAN ^10 -> CHAINED MULTIPLY (line ~1299)
   Replaced tmp^10 with t2=tmp*tmp, t4=t2*t2, t8=t4*t4, result=t8*t2.
   Replaced division by table dimensions with pre-computed inverse multipliers.
   At 100Hz with 6 balls: ~4,800 exponentiation ops/sec + ~2,400 divisions/sec eliminated.

7. VOL/BALLVEL ^2 -> x*x (line ~1319)
   Replaced ball.VelX^2 with vx*vx using cached locals.
   Cached COM reads for VelX/VelY into local vars.
   At 100Hz with 6 balls: ~1,200 exponentiation ops/sec eliminated.

8. ROLLINGTIMER REWRITE (line ~1346)
   - Set ball = BOT(b) caches ball object reference.
   - Inlined BallVel/Vol/Pitch: single Sqr + cached vx/vy per ball.
   - Computed bpan/bfade once, reused for rolling and drop sounds.
   - Used pre-built BallRollStr/BallDropStr arrays.
   - Cached ball.z, ball.VelZ into locals for drop sound check.
   - Fixed VBScript And short-circuit in VelZ/z check (nested If).
   - Cached UBound(BOT) once.
   At 100Hz with 6 balls: ~6,000 COM reads/sec + ~3,000 function calls/sec eliminated.

9. ONBALLBALLCOLLISION ^2 -> x*x (line ~1392)
   Cached Csng(velocity) into cv, used cv*cv in both branches.

10. BALLSHADOW_TIMER REWRITE (line ~1405)
    Cached UBound(BOT) once. Cached BOT(b).X and BOT(b).Z into bx/bz locals.
    Used bz for all height comparisons (was reading BOT(b).Z 3-4 times per ball).
    Added .visible delta guard (only writes when changed).
    Fixed VBScript And short-circuit in Z range check (nested If).
    At 60fps with 6 balls: ~2,160 COM reads/sec + ~360 redundant writes/sec eliminated.

11. PRIMT_TIMER COM CACHING (line ~1498)
    Cached l28a.State and l29a.State once each (was reading 3x each per tick).
    Pre-computed visibility value into local. Eliminates 4 redundant COM reads/tick.

ESTIMATED TOTAL SAVINGS
========================
Conservative total: ~29,000 redundant operations/sec eliminated.
Breakdown:
  - LampMod TypeName elimination: ~6,000 ops/sec
  - AudioFade/AudioPan ^10 elimination: ~4,800 ops/sec
  - RollingTimer inline + COM caching: ~9,000 ops/sec
  - UpdateFlipperLogo delta guards: ~1,200 ops/sec when at rest
  - BallShadow COM caching: ~2,520 ops/sec
  - LampTimer chglamp caching: ~2,000 ops/sec during lamp storms
  - String alloc elimination: ~1,800 allocs/sec
  - Pre-computed division constants: ~2,400 ops/sec
