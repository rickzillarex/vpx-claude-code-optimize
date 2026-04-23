Truck Stop (Bally 1988) v2.1.0 VBS — Optimization Changes
===============================================

1. FLIPPER TIMER 1ms -> 10ms (line ~1403)
   Changed RightFlipper.timerinterval from 1 to 10.
   100Hz is visually identical to 1000Hz for flipper correction.
   At 1000Hz: ~900 unnecessary timer calls/sec eliminated.

2. FLIPPER TIMER COM CACHING (line ~1406)
   Cached leftflipper/rightflipper .currentangle, .startangle, .endangle
   into local vars (lfCA, lfSA, lfEA, rfCA, rfSA, rfEA) at sub entry.
   Previously read each property 3-4 times per call.
   At 100Hz: ~1,200 COM reads/sec eliminated.

3. PRE-COMPUTED CONSTANTS (line ~1936)
   Added module-level vars: InvTWHalf (2/tablewidth), InvTHHalf (2/tableheight),
   TW_d2 (tablewidth/2), BS_d6 (BallSize/6).
   Eliminates repeated division in AudioFade, AudioPan, and shadow math.

4. PRE-BUILT STRING ARRAYS (line ~1936)
   Added BallRollStr(4), MetalRollStr(4), BallDropStr(4), FlasherMatStr(20).
   Eliminates "fx_ballrolling" & b, "fx_metalrolling" & b, "fx_balldrop" & b,
   and "Flashermaterial" & nr string concatenation in hot loops.
   At 100Hz with 4 balls: ~2,400 string allocs/sec eliminated.

5. AUDIOFADE/AUDIOPAN/PAN ^10 -> CHAINED MULTIPLY (line ~907)
   Replaced tmp^10 with t2=tmp*tmp, t4=t2*t2, t8=t4*t4, result=t8*t2.
   Also replaced division by table width/height with pre-computed inverse multipliers.
   At 100Hz with 4 balls x 2 calls each: ~3,200 exponentiation ops/sec eliminated.

6. VOL/BALLVEL/BALLSPEED ^2 -> x*x (line ~940)
   Replaced ball.VelX^2 + ball.VelY^2 with vx*vx + vy*vy using cached locals.
   Cached COM reads (ball.VelX, ball.VelY) into local vars.
   At 100Hz: ~800 exponentiation ops/sec + ~800 COM reads/sec eliminated.

7. ROLLINGTIMER REWRITE (line ~1089)
   - Set ball = BOT(b) caches ball object reference, eliminating repeated array indexing.
   - Inlined BallVel, Vol, Pitch computation: one Sqr + cached vx/vy per ball.
   - Computed bpan/bfade once, reused for both rolling and metal sound paths.
   - Used pre-built BallRollStr/MetalRollStr/BallDropStr arrays.
   - Cached ball.z, ball.VelZ into locals for drop sound check.
   - Used pre-computed TW_d2/BS_d6 for shadow positioning.
   - Added .Visible guard on BallShadow write.
   - Cached UBound(BOT) once.
   - Fixed VBScript And short-circuit trap in VelZ/z check (nested If).
   At 100Hz with 4 balls: ~4,800 COM reads/sec + ~2,400 function calls/sec eliminated.

8. FLASHFLASHER EXPONENTIATION (line ~1977)
   Replaced ^2.5, ^3, ^2 with chained multiplies (lvl2, lvl3, lvl2*Sqr(lvl)).
   Used pre-built FlasherMatStr(nr) and pre-computed RGBWhite.
   Cached FlashLevel(nr) into local lvl.
   At ~100Hz per flasher: ~500 exponentiation ops/sec eliminated per active flasher.

9. UPDATEFLIPPERS DELTA GUARDS (line ~1800)
   Added module-level lastLFAngle/lastRFAngle/lastLFTAngle/lastRFTAngle tracking.
   Cached CurrentAngle into local once per flipper.
   Only writes RotY to shadow and primitive when angle actually changed.
   At 60fps: when flippers at rest, ~480 COM writes/sec eliminated.

10. FLIPPERPOLARITY PROCESSBALLS COM CACHING (line ~1318)
    Cached Flipper.StartAngle, .CurrentAngle, .EndAngle into locals fSA/fCA/fEA.
    Eliminates 5 redundant COM reads per ProcessBalls call.

11. REMOVEBALL TypeName -> IsEmpty (line ~1287)
    Replaced TypeName(balls(x)) = "IBall" with Not IsEmpty(balls(x)).
    TypeName does COM reflection + string alloc; IsEmpty is a VarType check.

12. CORTRACKER UBOUND CACHING (line ~1903)
    Cached UBound(allballs) into ubAll. Used indexed For loop instead of For Each
    for second pass to avoid redundant iterator overhead.

ESTIMATED TOTAL SAVINGS
========================
Conservative total: ~16,000 redundant operations/sec eliminated.
Breakdown:
  - Flipper timer reduction (1000Hz->100Hz): ~900 calls/sec
  - Flipper COM caching: ~1,200 reads/sec
  - AudioFade/AudioPan ^10 elimination: ~3,200 ops/sec
  - RollingTimer inline + COM caching: ~7,200 ops/sec
  - FlashFlasher exponentiation: ~500 ops/sec per active flasher
  - UpdateFlippers delta guards: ~480 writes/sec when at rest
  - String alloc elimination: ~2,400 allocs/sec
  - Minor: BallVel/Vol/BallSpeed caching, TypeName removal
