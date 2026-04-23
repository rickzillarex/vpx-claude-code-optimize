AC-DC Premium 1.5 VBS — Optimization Changes
===============================================

1. PRE-COMPUTED CONSTANTS (lines ~425-450)
   Added module-level vars computed once at script load:
   - PIover180 (Pi/180): eliminates per-tick division in cannon, flipper, bell timers.
   - InvTWHalf (2/ACDC.Width), InvTHHalf (2/ACDC.Height): eliminates division in AudioFade/Pan.
   - HalfTableW (ACDC.Width/2), BallsizeD6 (Ballsize/6): eliminates division in BallShadowUpdate.
   - BallRollStr(0..6): pre-built "fx_ballrolling0" through "fx_ballrolling6".
   - gBOT: module-level variable for shared GetBalls result.
   - lastGIRed/lastGIBlue/lastGILowPF/lastGIWhite: GI delta-guard tracking (init -1).
   - lastFIRER/lastFIREG/lastFIREB: FIRE button delta-guard tracking (init -1).

2. PRE-COMPUTED BELL TRIG CONSTANTS (lines ~634-635)
   Added sin5deg = sin(5 * PIover180) and cos5deg = cos(5 * PIover180).
   BellMove_Timer uses these every tick for arm rotation. Eliminates 2 sin/cos calls per tick.
   At 100Hz: 200 trig calls/sec eliminated.

3. UPDATECANNON_TIMER — PIover180 + COM CACHING (line ~489)
   Replaced Pi/180 with PIover180 pre-computed constant.
   Cached laser rotation value into local var.
   Eliminates 1 division + redundant COM read per tick.
   At 100Hz: 100 divisions/sec + 100 COM reads/sec eliminated.

4. FLIPPER1_TIMER — COM CACHING (line ~569)
   Cached Flipper1.currentangle and Flipper2.currentangle into f1ca/f2ca locals.
   Flipper1.currentangle was read 6+ times per tick (3 comparisons + 3 assignments).
   Replaced Pi/180 with PIover180 for leg trig calculations.
   At 100Hz: ~1,200 COM reads/sec eliminated + 200 divisions/sec.

5. BELLMOVE_TIMER — PRE-COMPUTED TRIG + COM CACHING (line ~637)
   Replaced sin(5*PI/180)/cos(5*PI/180) with pre-computed sin5deg/cos5deg.
   Cached sin/cos arm computation into locals.
   Replaced Pi/180 with PIover180.
   At 100Hz: 400 trig calls/sec + 100 divisions/sec eliminated.

6. BELLK_TIMER — EXPONENTIATION + COM CACHING (line ~671)
   Replaced brakew^3 with bw3 = brakew*brakew*brakew (eliminates exponentiation).
   Cached Bell.RotX into bRot local variable.
   Replaced Pi/180 with PIover180.
   At 100Hz: 100 exponentiation ops/sec + 100 COM reads/sec + 100 divisions/sec eliminated.

7. LAMPTIMER_TIMER — 2D ARRAY CACHING (line ~846)
   Cached chgLamp(ii,0) and chgLamp(ii,1) into cIdx/cVal local variables at loop top.
   Each 2D array access is more expensive than a local var read in VBScript.
   Loop runs ~50 iterations per tick.
   At 100Hz: ~10,000 2D array lookups/sec eliminated.

8. FIRE BUTTON — DELTA-GUARDED RGB WRITES (within UpdateLamps)
   FIRE button sets 4 objects to the same RGB color every tick unconditionally.
   Added lastFIRER/G/B tracking. Only computes RGB() and writes when values change.
   FIRE color changes rarely (a few times per second) but UpdateLamps runs at ~62Hz.
   SAVES: ~248 RGB() calls/sec + ~992 COM writes/sec when color is unchanged.

9. UPDATEGIS — FULL DELTA-GUARD REWRITE (line ~1044)
   Each GI channel (Red, Blue, LowPF, White) now only updates when its LampState value changes.
   Previous: ran all 4 channel updates unconditionally every tick, each with For Each loops
   and repeated LampState()/50 division.
   Now: cached LampState/50 into local scale vars; each channel guarded by lastGIRed/Blue/LowPF/White.
   GI values change a few times per second; timer runs at ~62Hz.
   SAVES: ~3,600 COM writes/sec + ~248 divisions/sec when GI is stable (vast majority of time).

10. LAMPMOD — TYPENAME() ELIMINATION (line ~1213)
    Original LampMod used TypeName(object) to branch between Light and Flasher handling.
    TypeName() is COM reflection + string allocation + string comparison.
    Created LampModLight (line ~1224) and LampModFlasher (line ~1229) as typed variants.
    Updated ~44 calls in UpdateLamps to use the correct typed variant based on object type:
    - Flasher objects (f*): LampModFlasher
    - Light objects (f*r reflections, l151-l158 LED flames): LampModLight
    At ~62Hz: ~2,728 TypeName() calls/sec eliminated.

11. RGBLED — TYPENAME() ELIMINATION (line ~1191)
    Original RGBLED used TypeName(object) to branch between Light and Flasher.
    Created RGBLEDFlasher (line ~1197) for the 2 flasher calls (l40a, l39a).
    Remaining ~14 RGBLED calls are lights and keep the original function (TypeName removed,
    direct light path only).
    At ~62Hz: ~992 TypeName() calls/sec eliminated.

12. VOL/PAN/AUDIOFADE/BALLVEL — ^10/^2 ELIMINATION + PRE-COMPUTED INVERSES (lines ~1238-1280)
    Vol (line ~1238): BallVel(ball)^2 → bv*bv.
    Pan (line ~1243): tmp^10 → chained multiply (t2*t2→t4, t4*t4→t8, t8*t2).
      tableobj.x * 2/ACDC.Width → tableobj.x * InvTWHalf.
    AudioFade (line ~1256): tmp^10 → chained multiply.
      tableobj.y * 2/ACDC.Height → tableobj.y * InvTHHalf.
    BallVel (line ~1273): ball.VelX^2 + ball.VelY^2 → vx*vx + vy*vy with cached COM reads.
    Called from RollingSoundUpdate per ball per tick + sound event handlers.
    SAVES: ~4,200 exponentiation ops/sec + ~2,100 divisions/sec (at 6 balls × 100Hz).

13. REALTIMEUPDATES — SHARED GETBALLS (line ~1400)
    Both RollingSoundUpdate and BallShadowUpdate called GetBalls independently.
    Now calls GetBalls once into module-level gBOT, shared by both subs.
    GetBalls is a COM call that allocates an array of ball objects.
    SAVES: 1 GetBalls COM call per frame (~100Hz = 100 COM calls/sec + 100 array allocs/sec).

14. ROLLINGSOUNDUDATE — FULL REWRITE (line ~1423)
    Uses shared gBOT (eliminates GetBalls).
    Pre-built BallRollStr(b) replaces "fx_ballrolling" & b concatenation.
    Inlined BallVel: cached VelX/VelY into locals, single SQR call per ball.
    Inlined Vol/Pitch computation to avoid redundant BallVel re-calls.
    At 6 balls × 100Hz: ~600 string allocs/sec + ~1,200 COM reads/sec + ~600 function calls/sec eliminated.

15. BALLSHADOWUPDATE — COM CACHING + DELTA GUARDS (line ~1510)
    Uses shared gBOT (eliminates GetBalls).
    Cached BOT(b).X, BOT(b).Y, BOT(b).Z into bx/by/bz locals (each read once instead of 2-3 times).
    Replaced ACDC.Width/2 with pre-computed HalfTableW.
    Replaced Ballsize/6 with pre-computed BallsizeD6.
    Added delta-guarded .visible writes (only write when value changes).
    At 6 balls × 100Hz: ~3,600 COM reads/sec + ~1,200 divisions/sec eliminated + redundant visible writes.

16. ONBALLBALLCOLLISION — ^2 ELIMINATION (line ~1462)
    Replaced Csng(velocity)^2 with cv*cv (cached Csng result).
    Called on every ball-ball collision (sporadic but avoids exponentiation).

ESTIMATED TOTAL SAVINGS
========================
- TypeName() eliminations: ~3,720 calls/sec
- COM property reads eliminated: ~8,400/sec
- COM property writes eliminated: ~4,840/sec (delta guards)
- Exponentiation ops eliminated: ~4,400/sec
- Division ops eliminated: ~3,850/sec
- String allocations eliminated: ~600/sec
- GetBalls calls eliminated: 100/sec
- 2D array lookups eliminated: ~10,000/sec
- Trig calls eliminated: ~600/sec

Conservative total: ~36,500 redundant operations/sec eliminated.
