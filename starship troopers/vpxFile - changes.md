Starship Troopers vpxFile.vbs — Optimization Changes
===============================================

1. PRE-COMPUTED CONSTANTS (line ~956)
   Added module-level vars: InvTWHalf (2/tablewidth), InvTHHalf (2/tableheight),
   TW_d2 (tablewidth/2), BS_d6 (BallSize/6).
   Eliminates repeated division in AudioFade, AudioPan, and shadow math.

2. PRE-BUILT STRING ARRAYS (line ~956)
   Added BallRollStr(5), BallDropStr(5), FlasherMatStr(20), RGBWhite.
   Eliminates "fx_ballrolling" & b, "fx_ball_drop" & b, "Flashermaterial" & nr,
   and RGB(255,255,255) allocation in hot loops.
   At 100Hz with 5 balls: ~1,500 string allocs/sec eliminated.

3. LAMPTIMER INTERVAL 5ms -> 16ms (line ~619)
   Changed from 200Hz to 62Hz. LampTimer drives visual lamp fading which is
   visually indistinguishable at frame-rate (60fps). Adjusted FlashSpeedUp/Down
   proportionally (x3.2) to maintain identical fade timing.
   At 200Hz: ~138 unnecessary timer calls/sec eliminated.
   UpdateLamps has ~80 NFadeL calls each: ~11,000 function calls/sec eliminated.

4. LAMPTIMER chglamp CACHING (line ~624)
   Cached chgLamp(ii,0) and chgLamp(ii,1) into cIdx/cVal locals at loop top.
   Eliminates 2 redundant 2D array dereferences per lamp per tick.

5. FLIPPERTIMER DELTA GUARDS (line ~1149)
   Added lastLFAngle/lastRFAngle/lastMFAngle tracking.
   Cached CurrentAngle into locals once per flipper (was read 2x each).
   Only writes RotZ/RotY when angle changed. At rest: skip all COM writes.
   At 60fps: ~360 COM reads + ~360 COM writes/sec eliminated when at rest.

6. FLASHFLASHER EXPONENTIATION (line ~1004)
   Replaced ^2.5 with lvl2*Sqr(lvl), ^3 with lvl2*lvl, ^2 with lvl*lvl.
   Cached FlashLevelb(nr) into local lvl. Used FlasherMatStr(nr) and RGBWhite.
   Fixed VBScript And short-circuit trap (split compound conditions).
   16 flashers x ~100Hz each: ~4,800 exponentiation ops/sec eliminated.

7. BALLSHADOWUPDATE REWRITE (line ~1164)
   Cached UBound(BOT) once. Cached BOT(b).X and BOT(b).Z into bx/bz locals.
   Used pre-computed TW_d2/BS_d6 instead of Table1.Width/2 and BallSize/6.
   Added .visible delta guard (only writes when changed).
   At 60fps with 5 balls: ~1,500 COM reads/sec + ~300 redundant writes/sec eliminated.

8. AUDIOFADE/AUDIOPAN/PAN ^10 -> CHAINED MULTIPLY (line ~1272)
   Replaced tmp^10 with t2=tmp*tmp, t4=t2*t2, t8=t4*t4, result=t8*t2.
   Replaced division by table dimensions with pre-computed inverse multipliers.
   At 100Hz with 5 balls: ~4,000 exponentiation ops/sec + ~2,000 divisions/sec eliminated.

9. VOL/BALLVEL/VOLZ ^2 -> x*x (line ~1305)
   Replaced ball.VelX^2 with vx*vx using cached locals.
   Cached COM reads for VelX/VelY into local vars.
   At 100Hz with 5 balls: ~1,000 exponentiation ops/sec eliminated.

10. ROLLINGTIMER REWRITE (line ~1360)
    - Set ball = BOT(b) caches ball object reference.
    - Inlined BallVel/Vol/Pitch: single Sqr + cached vx/vy per ball.
    - Computed bpan/bfade once, reused for PlaySound.
    - Used pre-built BallRollStr/BallDropStr arrays.
    - Cached ball.z and ball.VelZ into locals.
    - Fixed VBScript And short-circuit in VelZ/z check.
    - Cached UBound(BOT) once.
    At 100Hz with 5 balls: ~5,000 COM reads/sec + ~2,500 function calls/sec eliminated.

11. ONBALLBALLCOLLISION ^2 -> x*x (line ~1401)
    Cached Csng(velocity) into cv, used cv*cv.

12. DEBUG.PRINT REMOVAL (line ~597)
    Commented out Debug.Print in BBUGtimer_Timer.
    Debug.Print allocates COM strings even when no debugger is attached.

ESTIMATED TOTAL SAVINGS
========================
Conservative total: ~29,000 redundant operations/sec eliminated.
Breakdown:
  - LampTimer frequency reduction (200Hz->62Hz): ~11,000 function calls/sec
  - AudioFade/AudioPan ^10 elimination: ~4,000 ops/sec
  - FlashFlasher exponentiation (16 flashers): ~4,800 ops/sec
  - RollingTimer inline + COM caching: ~7,500 ops/sec
  - FlipperTimer delta guards: ~720 ops/sec when at rest
  - BallShadow COM caching: ~1,800 ops/sec
  - String alloc elimination: ~1,500 allocs/sec
