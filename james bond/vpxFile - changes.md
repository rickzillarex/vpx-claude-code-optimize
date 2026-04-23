James Bond VBS — Optimization Changes
===============================================

1. PRE-COMPUTED INVERSE MULTIPLIERS (line ~812)
   Added InvTWHalf = 2/TableWidth and InvTHHalf = 2/TableHeight at module level.
   Pan() and AudioFade() now multiply instead of dividing per call.
   At ~60Hz with 5 balls: 600 divisions/sec eliminated.

2. EXPONENTIATION ELIMINATION — ^10 in Pan/AudioFade (line ~828, ~850)
   Replaced ^10 with chained multiply: t2=tmp*tmp, t4=t2*t2, t8=t4*t4, result=t8*t2.
   Pan and AudioFade are called 2x per ball per frame from RollingUpdate + drop sounds.
   At 60Hz with 5 balls: ~600 Exp(10*Log(x)) ops/sec eliminated.

3. EXPONENTIATION ELIMINATION — ^2 in BallVel, Vol (line ~821, ~843)
   BallVel: ball.VelX^2 + ball.VelY^2 replaced with vx*vx + vy*vy.
   Vol: BallVel(ball)^2 replaced with bv*bv.
   COM reads in BallVel cached into locals (vx, vy).
   At 60Hz with 5 balls: ~600 exponentiation ops + ~600 COM reads/sec eliminated.

4. PRE-COMPUTED BS_d2 CONSTANT (line ~822)
   Added BS_d2 = BallSize/2 at module level. Replaces per-frame division
   in RollingUpdate shadow height and ball creation calls.
   At 60Hz with 5 balls: 300 divisions/sec eliminated.

5. PRE-BUILT BallRollStr ARRAY (line ~878)
   Built "fx_ballrolling0" through "fx_ballrolling19" at init.
   Eliminates "fx_ballrolling" & b string concatenation in RollingUpdate.
   At 60Hz with 5 balls: ~900 string allocs/sec eliminated (3 uses per ball).

6. RollingUpdate REWRITE — COM PROPERTY CACHING (line ~895)
   Cache BOT(b) as local 'ball' via Set. Cache ball.X/Y/Z/VelX/VelY into
   bx/by/bz/bvx/bvy locals. Cache UBound(BOT) once at sub entry.
   Eliminates repeated COM property reads per ball per frame.
   Before: ~20+ COM reads per ball per iteration.
   After: 6 COM reads per ball per iteration.
   At 60Hz with 5 balls: ~4,200 COM reads/sec eliminated.

7. RollingUpdate REWRITE — INLINED BallVel/Vol/Pitch (line ~895)
   BallVel, Vol, Pitch were called separately, each re-reading COM properties.
   Now computed inline: bv = SQR(bvx*bvx + bvy*bvy), ballvol/ballpitch
   computed directly from bv. Eliminates 3 redundant function calls + 6
   redundant COM reads per ball per frame.
   At 60Hz with 5 balls: ~1,800 function calls + ~1,800 COM reads/sec eliminated.

8. RollingUpdate — CACHED VelZ FOR DROP SOUNDS (line ~931)
   VelZ and z reads for drop sound check now use cached bvz/bz locals.
   Nested If replaces non-short-circuiting And chain (VBScript evaluates both sides).
   At 60Hz with 5 balls: ~600 COM reads/sec eliminated.

9. GIUpdateTimer — ELIMINATED GetBalls (line ~697)
   Replaced GetBalls COM call with BallsOnPlayfield variable check.
   The variable is already maintained by CreateNewBall/BallDrained logic.
   At timer frequency (~10Hz): 10 GetBalls COM allocs/sec eliminated.

10. RealTime_Timer — FLIPPER ANGLE WRITE GUARDS (line ~2559)
    Added lastLFAngle, lastLF001Angle, lastRFAngle tracking variables.
    RotZ writes guarded: only write when CurrentAngle actually changes.
    During flipper rest (majority of gameplay): 3 writes become 0 per tick.
    At 60Hz: ~180 COM writes/sec eliminated when flippers at rest.

11. RainbowTimer — PRE-COMPUTED RGB COLORS (line ~2800)
    RGB() computed once per tick into cDim/cFull locals, then applied to all lights.
    Previously computed RGB() per light per tick (2 RGB calls per light).
    With ~10 rainbow lights at ~30Hz: ~600 RGB() calls/sec reduced to ~60.

12. OnBallBallCollision — ^2 ELIMINATION (line ~950)
    velocity^2 replaced with velocity*velocity.
    Occasional (per collision), minor savings.

ESTIMATED TOTAL SAVINGS
========================
- COM property reads: ~7,200/sec (RollingUpdate caching + inlining + GIUpdateTimer)
- Exponentiation: ~1,200/sec (Pan/AudioFade ^10, BallVel/Vol ^2)
- Division: ~900/sec (InvTWHalf/InvTHHalf/BS_d2 pre-computation)
- String allocation: ~900/sec (BallRollStr pre-built array)
- COM writes: ~180/sec (flipper RotZ guards at rest)
- Function calls: ~1,800/sec (inlined BallVel/Vol/Pitch in RollingUpdate)
- RGB() calls: ~540/sec (RainbowTimer pre-computation)
- GetBalls alloc: ~10/sec (GIUpdateTimer)
Conservative total: ~12,700 redundant operations/sec eliminated.
