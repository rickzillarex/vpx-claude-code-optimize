JP's Avengers LE VBS — Optimization Changes
=============================================

1. PRE-COMPUTED CONSTANTS (line ~602-616)
   Added module-level vars:
   - InvTWHalf (2/TableWidth), InvTHHalf (2/TableHeight): eliminates division per Pan/AudioFade call.
   - BallRollStr(0..19): pre-built "fx_ballrolling0" through "fx_ballrolling19", eliminates string concat.
   - PIover180 (Pi/180): eliminates division per SpinKicker_Timer and collision trig calls.
   - lastLFAngle, lastRFAngle: flipper _Animate delta-write tracking.

2. FLIPPER TIMER 1ms → 10ms (line ~529)
   Changed LeftFlipper.TimerInterval from 1 to 10 (1000Hz → 100Hz).
   LeftFlipper_Timer handles BOTH flippers' FlipperTricks (SOS + EOS + LiveCatch).
   At 1ms: 1000 calls/sec with ~12 COM reads per call = 12,000 COM reads/sec.
   At 10ms: 100 calls/sec = 1,200 COM reads/sec.
   SAVES: ~10,800 COM reads/sec. 100Hz is visually identical to 1000Hz for flipper correction.
   THIS IS THE #1 FIX FOR THIS TABLE.

3. LEFTFLIPPER_TIMER — CACHE COM PROPERTIES (line ~532)
   Cached LeftFlipper.CurrentAngle/StartAngle into Lca/Lsa locals.
   Cached RightFlipper.CurrentAngle/StartAngle into Rca/Rsa locals.
   Before: CurrentAngle read 2× per flipper (SOS check + EOS check), StartAngle read 1×.
   After: 1 COM read each, rest are local var access.
   Even at 100Hz (after Fix 2): saves ~400 COM reads/sec.

4. VOL — ELIMINATE ^2 (line ~618)
   Replaced `BallVel(ball) ^2` with `bv * bv` (cached BallVel result).
   SAVES: 1 exponentiation op per call.

5. PAN — ELIMINATE ^10, USE PRE-COMPUTED INVERSE (line ~622)
   Replaced `tmp ^10` with chained multiply: t2=tmp*tmp, t4=t2*t2, t8=t4*t4, result=t8*t2.
   Replaced `ball.x * 2 / TableWidth` with `ball.x * InvTWHalf`.
   SAVES: ~1,400 exponentiation ops/sec + ~1,400 divisions/sec (7 balls × 100Hz × 2 calls).

6. BALLVEL — CACHE COM, ELIMINATE ^2 (line ~636)
   Replaced `ball.VelX ^2 + ball.VelY ^2` with cached locals: vx*vx + vy*vy.
   SAVES: 2 COM reads and 2 exponentiation ops per call.

7. AUDIOFADE — ELIMINATE ^10, USE PRE-COMPUTED INVERSE (line ~642)
   Same treatment as Pan: chained multiply + InvTHHalf.
   SAVES: ~1,400 exponentiation ops/sec + ~1,400 divisions/sec.

8. ROLLINGTIMER_TIMER REWRITE (line ~681)
   - Cache UBound(BOT) into ubBot local.
   - Use pre-built BallRollStr(b) instead of "fx_ballrolling" & b (3 sites).
   - Inline BallVel: cache BOT(b).VelX/VelY into bvx/bvy, compute bv = SQR(bvx*bvx+bvy*bvy).
   - Inline Vol: use Csng(bv*bv/2000) directly (eliminates BallVel→Vol double call chain).
   - Inline Pitch: use bv*20 directly (eliminates BallVel→Pitch double call chain).
   - Cache BOT(b).z into bz, reuse for rolling check + drop sound check.
   - Cache BOT(b).VelZ into bvz for drop sound check.
   - Speed control: reuse cached bvx/bvy for initial VelX/VelY reads.
   SAVES per ball per tick:
   - 8 COM reads eliminated (VelX/VelY/z/VelZ read once each instead of 3-4× via helper chains)
   - 2 Sqr() calls eliminated (BallVel called 3× before: Vol, Pitch, direct; now computed once)
   - 1 string alloc eliminated
   At 19 balls × ~100Hz: up to ~15,200 COM reads/sec + ~3,800 Sqr/sec + 1,900 string allocs/sec.

9. FLIPPER _ANIMATE — DELTA-GUARDED WRITES (line ~738)
   LeftFlipper_Animate and RightFlipper_Animate now cache CurrentAngle and only
   write RotZ when the angle has changed from the last frame.
   _Animate fires every frame (~60Hz). When flippers are at rest (~80% of gameplay):
   SAVES: ~120 redundant COM writes/sec (2 flippers × 60fps).

10. SPINKICKER_TIMER — CACHE TRIG + COM READS (line ~781)
    - Use pre-computed PIover180 instead of (PI/180) per tick.
    - Cache cos/sin results: compute sx0/sy0 once, derive sx1/sy1 by symmetry.
      Before: SpBall(0).x/.y set, then SpBall(1) re-reads SpBall(0).x/.y from COM.
      After: sx1/sy1 computed from locals, zero re-reads.
    - Cache CubeB.RotZ into cRot local: read once, write back once after modification.
      Before: CubeB.RotZ read 3× and written 4× (CubeB + CubePr + CubePr2 + CubeBase).
      After: read 1×, written 4× (writes are necessary — different objects).
    - LampPost positions use cached sx0/sy0/sx1/sy1 instead of re-reading SpBall COM.
    SAVES: ~6 COM reads/tick. At ~60Hz (timer freq): ~360 COM reads/sec.

11. ONBALLBALLCOLLISION — ELIMINATE ^2 (line ~842)
    Replaced `Csng(velocity) ^2` with `cv * cv` in the non-spinner collision branch.
    Infrequent, but free to fix.

12. COLLISION TRIG — PRE-COMPUTED PIover180 (line ~849)
    Replaced `RotAdj * Pi / 180` with `RotAdj * PIover180` in spinner collision calc.
    Eliminates 1 division per ball-spinner collision.


ESTIMATED TOTAL SAVINGS
========================
| Category                  | Before (ops/sec) | After (ops/sec) | Saved     |
|---------------------------|-------------------|-----------------|-----------|
| Flipper timer calls       | 1,000             | 100             | 900       |
| Flipper COM reads         | 12,000            | 1,200           | ~10,800   |
| Exponentiation (^10)      | 2,800             | 0               | 2,800     |
| Exponentiation (^2)       | ~3,000            | 0               | ~3,000    |
| Division (audio/trig)     | ~2,860            | ~60             | ~2,800    |
| String alloc (rolling)    | ~1,900            | 0               | ~1,900    |
| COM reads (rolling)       | ~15,200           | ~3,800          | ~11,400   |
| COM writes (flippers)     | ~120              | ~0 (at rest)    | ~120      |
| COM reads (spinkicker)    | ~360              | ~0              | ~360      |
| Sqr calls (rolling)       | ~3,800            | ~1,900          | ~1,900    |
| TOTAL                     |                   |                 | ~36,000   |


NOT CHANGED (and why)
======================
- body_Animate, arms_Animate, gate*_Animate: simple 1-2 COM writes per frame.
  Could add delta guards but savings are marginal (~360 COM writes/sec total).
- GIUpdate: fires on GI state change only (event-driven, not per-frame).
- cvpmVLock2 class: entirely event-driven.
- SolCallback handlers: fire on solenoid events only.
- All _Hit event handlers: fire a few times/sec at most.
- Table1_Init: runs once.
- FlashForMs-style flasher modulation subs (Flasher18-32): per-solenoid-event,
  already minimal (1 division + 1 COM write each).
