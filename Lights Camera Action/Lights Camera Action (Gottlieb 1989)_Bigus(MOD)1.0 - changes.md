Lights Camera Action (Gottlieb 1989) VBS — Optimization Changes
===============================================

1. AUDIOFADE/AUDIOPAN/PAN ^10 → CHAINED MULTIPLY (lines ~378-409)
   Replaced `tmp ^10` with t2=tmp*tmp, t4=t2*t2, t8=t4*t4, result=t8*t2
   (4 multiplies instead of Exp(10*Log(x))). Applied to all three functions
   (AudioFade, AudioPan, Pan) and both positive/negative branches.
   At 100Hz with 5 balls (3 funcs x 2 calls each): ~3,000 exponent ops/sec eliminated.

2. AUDIOFADE/AUDIOPAN/PAN DIVISION → PRE-COMPUTED INVERSE (lines ~378-409)
   Replaced `tableobj.y * 2 / table1.height` with `tableobj.y * InvTHHalf`
   and `tableobj.x * 2 / table1.width` with `tableobj.x * InvTWHalf`.
   InvTHHalf/InvTWHalf computed once in Table1_Init.
   At 100Hz with 5 balls: ~3,000 divisions/sec eliminated.

3. BALLVEL ^2 → x*x + COM CACHING (line ~432)
   Replaced `ball.VelX ^2 + ball.VelY ^2` with cached locals
   `vx*vx + vy*vy`. Eliminates 2 exponentiation ops and caches COM reads.
   At 100Hz with 5 balls: ~1,000 exponent ops + ~1,000 COM reads eliminated.

4. VOL/VOLMULTI/DVOLMULTI/BALLROLLVOL ^2 → x*x (lines ~411-427)
   All volume functions now cache BallVel result and use `bv * bv` instead
   of `BallVel(ball) ^2`. Eliminates double BallVel call (was called 2x).
   At 100Hz with 5 balls: ~500 exponent ops + ~500 function calls eliminated.

5. VOLZ ^2 → x*x (line ~441)
   Replaced `BallVelZ(ball) ^2` with `bvz * bvz`.

6. DEBUG.PRINT REMOVED FROM DVolMulti (line ~419)
   Removed `debug.print DVolMulti` from DVolMulti function.
   debug.print does COM string allocation + output write per call.
   Called from sound functions: ~200 COM allocs/sec eliminated.

7. PRE-BUILT BALL ROLL/DROP STRING ARRAYS (module level)
   Pre-built BallRollStr(0..5) = "fx_ballrolling0".."fx_ballrolling5"
   and BallDropStr(0..5) = "fx_ball_drop0".."fx_ball_drop5" at init.
   Replaces per-tick string concatenation `"fx_ballrolling" & b`.
   At 100Hz with 5 balls: ~1,500 string allocs/sec eliminated.

8. ROLLINGTIMER INLINED BallVel/Vol/Pitch (line ~475)
   Inlined BallVel, Vol, and Pitch into the per-ball loop body.
   Cached ball.VelX/VelY/VelZ/z into locals (4 COM reads vs 10+).
   Set ball = BOT(b) caches the ball object reference.
   Eliminated redundant BallVel calls (was called 3x per ball via
   BallVel, Vol->BallVel, Pitch->BallVel).
   At 100Hz with 5 balls: ~3,000 COM reads/sec + ~1,000 Sqr/sec eliminated.

9. FLIPPERS_TIMER COM CACHING + DELTA GUARDS (line ~315)
   Cached LeftFlipper.CurrentAngle, Flipper1.CurrentAngle,
   RightFlipper.CurrentAngle into locals. Added delta guards: only
   write objRotZ when angle has changed from last tick.
   When flippers at rest: eliminates 3 COM reads + 3 COM writes per tick.
   At timer Hz: ~600 COM ops/sec eliminated when flippers idle.

10. REALTIME_TIMER DELTA GUARDS (line ~555)
    Cached CurrentAngle for flipper shadows. Added delta guard: only write
    RotZ when angle actually changed.
    At 60fps: ~240 COM writes/sec eliminated when flippers idle.

11. TARGETTIMER_TIMER DELTA GUARDS (line ~330)
    Added lastTgtVis() tracking array. Reads 9 light states per tick but
    only writes Visible/Amount when the value has changed. During stable
    gameplay most flashers don't change every tick.
    At timer Hz (~100Hz): up to ~1,800 COM writes/sec eliminated when
    flasher states are stable.

12. BALLSHADOWUPDATE COM CACHING + VISIBLE GUARDS (line ~519)
    Cached BOT(b).X/Y/Z into locals (3 COM reads vs 5-7 per ball).
    Guard .visible writes with delta check.
    At timer Hz with 5 balls: ~500 COM reads + ~500 COM writes eliminated.

13. FLEXDMD UpdateFlexChar: ELIMINATED NewImage ALLOCATION (line ~855)
    Original code called `FlexDMD.NewImage("", FlexDMDDict.Item(value)).Bitmap`
    per LED update, allocating a new FlexDMD image object on every call.
    Now uses FlexDMDImageCache: pre-built image objects resolved from
    the dictionary at init. Per-update path does one Dictionary.Exists() +
    .Item() (cached image) instead of NewImage allocation.
    At LED update rate (~40 segments x ~30Hz): ~1,200 object allocs/sec eliminated.

14. FLEXDMD PRE-BUILT SEGMENT NAME STRINGS (module level)
    Pre-built SegStr(0..39) = "Seg0".."Seg39" at init.
    Eliminates `"Seg" & id` and `"Seg" & i` string concatenation in
    FlexDMD_Init and UpdateFlexChar.
    At LED update rate: ~1,200 string allocs/sec eliminated.

15. ONBALLBALLCOLLISION velocity^2 → v*v (line ~515)
    Replaced `Csng(velocity) ^2` with `velocity * velocity`.
    Event handler (~15/sec max) but free to fix.

16. PRE-COMPUTED CONSTANTS AT MODULE LEVEL
    Added InvTWHalf, InvTHHalf as pre-computed values in Table1_Init.
    Division by table dimensions replaced with multiplication by inverse
    throughout hot paths.


ESTIMATED TOTAL SAVINGS
========================
- Exponentiation elimination: ~4,500/sec (^10, ^2 across 3 audio funcs + vol funcs)
- COM read elimination:        ~5,000/sec (rolling, shadows, flippers)
- COM write guards:            ~3,100/sec (flippers, targets, shadows)
- String alloc elimination:    ~2,700/sec (ball roll strings, segment names)
- Object alloc elimination:    ~1,200/sec (FlexDMD NewImage per LED update)
- Division elimination:        ~3,000/sec (table dimension divisions)
- debug.print removal:         ~200/sec

Conservative total: ~19,700 redundant operations/sec eliminated.
