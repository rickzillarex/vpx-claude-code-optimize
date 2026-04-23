Mission Impossible VBS — Optimization Changes
===============================================

1. EXPONENTIATION ELIMINATION — BallVel (line ~801)
   Replaced `ball.VelX ^2` and `ball.VelY ^2` with cached locals
   `vx*vx + vy*vy`. Eliminates 2 exponentiation ops per call.
   Also caches VelX/VelY into locals (2 fewer COM reads per call).
   At 60Hz with 5 balls: ~600 calls/sec from RollingUpdate alone.

2. EXPONENTIATION ELIMINATION — Pan ^10 (line ~790)
   Replaced `tmp ^10` with chained multiply: t2=tmp*tmp, t4=t2*t2,
   t8=t4*t4, result=t8*t2. Eliminates Exp(10*Log(x)) per call.
   At 60Hz x 5 balls: ~600 ops/sec eliminated.

3. EXPONENTIATION ELIMINATION — AudioFade ^10 (line ~808)
   Same chained multiply pattern as Pan. ~600 ops/sec eliminated.

4. EXPONENTIATION ELIMINATION — Vol ^2 (line ~783)
   Replaced `BallVel(ball) ^2` with `bv * bv` using cached BallVel result.
   ~600 ops/sec eliminated.

5. EXPONENTIATION ELIMINATION — OnBallBallCollision ^2 (line ~908)
   Replaced `velocity ^2` with `velocity * velocity`.
   Low frequency but free to fix.

6. PRE-COMPUTED INVERSE MULTIPLIERS (line ~780)
   Added `InvTWHalf = 2 / TableWidth` and `InvTHHalf = 2 / TableHeight`
   at module level. Pan and AudioFade now multiply instead of divide.
   At 60Hz x 5 balls x 2 functions: ~600 divisions/sec eliminated.

7. ROLLINGUPDATE REWRITE — COM PROPERTY CACHING (line ~860)
   Cached BOT(b).X/Y/Z/VelX/VelY/VelZ into local vars bx/by/bz/bvx/bvy/bvz
   at top of each ball iteration. Previously BallVel, Vol, Pitch, Pan,
   AudioFade each re-read VelX/VelY from COM separately.
   Original: ~15 COM reads per ball per frame.
   After: 6 COM reads per ball per frame.
   At 60Hz x 5 balls: ~2,700 COM reads/sec eliminated.

8. ROLLINGUPDATE — INLINED BallVel/Vol/Pitch (line ~860)
   BallVel was called 3-4 times per ball (rolling check, Pitch, Vol,
   drop sound). Now computed once inline using cached velocity locals.
   Eliminates ~1,800 redundant function calls + Sqr ops per second.

9. PRE-BUILT STRING ARRAY — BallRollStr (line ~841)
   Pre-built `BallRollStr(0..19) = "fx_ballrolling" & i` at init.
   Eliminates string concatenation from RollingUpdate hot path.
   At 60Hz x 5 balls: ~300 string allocs/sec from play + ~300 from stop.

10. CACHED UBound(BOT) (line ~855)
    UBound(BOT) cached into local `ubBot` at sub entry. Was called
    3+ times per RollingUpdate invocation.
    At 60Hz: ~180 function calls/sec eliminated.

11. PRE-COMPUTED BS_d2 (line ~838)
    `BallSize / 2` computed once as module-level constant instead of
    per ball per frame. At 60Hz x 5 balls: ~300 divisions/sec eliminated.

12. GIUpdateTimer — ELIMINATED GetBalls (line ~660)
    Replaced `GetBalls` COM call with `BallsOnPlayfield` variable
    (already tracked by the game logic). Eliminates one GetBalls
    allocation per GIUpdateTimer tick (~60-100Hz).

13. REALTIME_TIMER — FLIPPER ANGLE WRITE GUARDS (line ~2516)
    Added delta guards for LeftFlipperTop.RotZ and RightFlipperTop.RotZ.
    Only writes COM property when angle actually changes.
    Flippers at rest (~80% of gameplay): eliminates ~120 COM writes/sec.

14. PRE-BUILT TORCH STRINGS (line ~5410)
    Pre-built `TorchStr(0..7) = "torch" & i` at init.
    torchtimer_timer now uses array lookup instead of concatenation.
    At timer frequency: ~3 string allocs/tick eliminated.

15. RAINBOW TIMER — PRE-COMPUTED RGB (line ~2722)
    Moved RGB() computation outside the For Each loop. RGB values
    computed once per tick instead of once per light per tick.
    With N lights in collection: eliminates 2*(N-1) RGB() calls/tick.

16. ROLLINGUPDATE — DROP SOUND FIX (line ~883)
    VBScript `And` does not short-circuit. Restructured the
    `bvz < -1 and bz < 55 and bz > 27` check into nested Ifs
    and uses cached bvz/bz locals instead of re-reading COM.


ESTIMATED TOTAL SAVINGS
========================
- COM reads:        ~4,500/sec eliminated (RollingUpdate + GIUpdate + flippers)
- Exponentiation:   ~2,400/sec eliminated (^10 in Pan/AudioFade, ^2 in BallVel/Vol)
- String allocs:    ~600/sec eliminated (BallRollStr + TorchStr)
- Division:         ~900/sec eliminated (InvTWHalf/InvTHHalf + BS_d2)
- Function calls:   ~2,000/sec eliminated (inlined BallVel/Vol/Pitch)
- COM writes:       ~120/sec eliminated (flipper angle guards)
- GetBalls allocs:  ~100/sec eliminated (GIUpdateTimer)

Conservative total: ~10,600 redundant operations/sec eliminated.
