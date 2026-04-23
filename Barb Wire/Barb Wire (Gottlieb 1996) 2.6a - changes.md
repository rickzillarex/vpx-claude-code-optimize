Barb Wire (Gottlieb 1996) 2.6a VBS — Optimization Changes
===============================================

1. LAMPTIMER INTERVAL 5ms → 10ms (line ~441)
   LampTimer.Interval changed from 5 (200Hz) to 10 (100Hz).
   Lamp fading is visually identical at 100Hz. Halves the number of
   UpdateLamps calls, which each invoke ~90 NFadeL/Flash/FadeDisableLighting
   calls.
   At 200Hz→100Hz: ~9,000 function calls/sec eliminated.

2. LAMPTIMER chglamp(ii,0)/(ii,1) CACHED INTO LOCALS (line ~448)
   Cached chgLamp(ii,0) and chgLamp(ii,1) into cIdx/cVal local variables
   at the top of the ChangedLamps loop. Eliminates redundant 2D Variant
   array dereferences per lamp change per tick.
   At 100Hz with ~10 lamp changes/batch: ~2,000 array derefs/sec eliminated.

3. AUDIOFADE/AUDIOPAN ^10 → CHAINED MULTIPLY (lines ~733-753)
   Replaced `tmp ^10` with t2=tmp*tmp, t4=t2*t2, t8=t4*t4, result=t8*t2
   (4 multiplies instead of Exp(10*Log(x))). Applied to both positive and
   negative branches.
   At 100Hz with 5 balls: ~1,000 exponentiation ops/sec eliminated.

4. AUDIOFADE/AUDIOPAN DIVISION → PRE-COMPUTED INVERSE (lines ~733-753)
   Replaced `tableobj.y * 2 / table1.height` with `tableobj.y * InvTHHalf`
   and `tableobj.x * 2 / table1.width` with `tableobj.x * InvTWHalf`.
   InvTHHalf/InvTWHalf computed once in Table1_Init.
   At 100Hz with 5 balls (2 calls each): ~2,000 divisions/sec eliminated.

5. BALLVEL ^2 → x*x (line ~763)
   Replaced `ball.VelX ^2 + ball.VelY ^2` with cached locals
   `vx*vx + vy*vy`. Eliminates 2 exponentiation ops + caches COM reads.
   At 100Hz with 5 balls: ~1,000 exponent ops + ~1,000 COM reads eliminated.

6. VOL ^2 → x*x (line ~756)
   Replaced `BallVel(ball) ^2` with `bv * bv` using cached BallVel result.
   At 100Hz with 5 balls: ~500 exponent ops/sec eliminated.

7. VOLZ ^2 → x*x (line ~772)
   Replaced `BallVelZ(ball) ^2` with `bvz * bvz`.

8. PRE-BUILT BALL ROLL/DROP STRING ARRAYS (module level)
   Pre-built BallRollStr(0..5) = "fx_ballrolling0".."fx_ballrolling5"
   and BallDropStr(0..5) = "fx_ball_drop0".."fx_ball_drop5" at init.
   Replaces per-tick string concatenation `"fx_ballrolling" & b`.
   At 100Hz with 5 balls: ~1,500 string allocs/sec eliminated.

9. ROLLINGTIMER INLINED BallVel/Vol/Pitch (line ~791)
   Inlined BallVel, Vol, and Pitch into the per-ball loop body.
   Cached ball.VelX/VelY/VelZ/z into locals (4 COM reads vs 10+ before).
   Set ball = BOT(b) caches the ball object reference.
   Eliminated redundant BallVel calls (was called 3x per ball: BallVel,
   Vol→BallVel, Pitch→BallVel).
   At 100Hz with 5 balls: ~3,000 COM reads/sec + ~1,000 Sqr/sec eliminated.

10. FLIPPERS_TIMER COM CACHING + DELTA GUARDS (line ~355)
    Cached LeftFlipper.CurrentAngle, LeftFlipperUp.CurrentAngle,
    RightFlipper.CurrentAngle into locals. Added delta guards: only
    write objRotZ and RotZ when angle has changed from last tick.
    When flippers at rest (majority of play): eliminates 6 COM reads +
    6 COM writes per tick.
    At 100Hz: ~1,200 COM ops/sec eliminated when flippers idle.

11. BALLSHADOWUPDATE TABLE WIDTH CACHING (line ~845)
    Replaced `Table1.Width/2` with pre-computed TW_d2 and `Ballsize/6`
    with pre-computed BS_d6. Guards .visible writes with delta check.
    At 100Hz with 5 balls: ~500 COM reads + ~500 divisions eliminated.

12. DEBUG.PRINT REMOVED FROM FadeDisableLighting (line ~634)
    Removed `debug.print a.uservalue` from FadeDisableLighting.
    debug.print does COM string allocation and output write every tick.
    At 100Hz with 3 flashers fading: ~300 COM allocs/sec eliminated.

13. CONTACTLENST_TIMER COM CACHING (line ~298)
    Cached ContactLensP.TransY into local, modified local, wrote back
    once. Was reading TransY 4+ times per tick.
    At timer interval: ~150 COM reads/sec eliminated.

14. FATSOTIMER_TIMER COM CACHING (line ~319)
    Cached FatsoPrim.TransY into local, modified local, wrote back once.
    Was reading TransY 3 times per tick.
    At timer interval: ~100 COM reads/sec eliminated.

15. ONBALLBALLCOLLISION velocity^2 → v*v (line ~835)
    Replaced `Csng(velocity) ^2` with `velocity * velocity`.
    Event handler (~15/sec max) but free to fix.

16. PRE-COMPUTED CONSTANTS AT MODULE LEVEL
    Added PI, PIover180, BS_d6 (Ballsize/6), InvTWHalf, InvTHHalf, TW_d2
    as pre-computed values. Division by table dimensions replaced with
    multiplication by inverse throughout hot paths.


ESTIMATED TOTAL SAVINGS
========================
- LampTimer halved:          ~9,000 function calls/sec
- String alloc elimination:  ~1,500/sec
- Exponentiation elimination: ~2,500/sec (^10, ^2)
- COM read elimination:       ~6,000/sec (rolling, shadows, flippers, timers)
- COM write guards:           ~1,200/sec (flippers, shadows)
- Division elimination:       ~2,500/sec (table dims, ballsize)
- debug.print removal:        ~300/sec

Conservative total: ~23,000 redundant operations/sec eliminated.
