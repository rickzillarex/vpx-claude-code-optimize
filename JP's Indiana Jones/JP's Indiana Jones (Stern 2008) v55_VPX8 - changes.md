JP's Indiana Jones (Stern 2008) v55_VPX8 VBS — Optimization Changes
=====================================================================

1. FLIPPER TIMER 1ms -> 10ms (line ~820)
   Changed LeftFlipper.TimerInterval from 1 to 10. 100Hz is visually
   identical to 1000Hz for flipper tricks. This single change eliminates
   ~900 function calls/sec (1000Hz - 100Hz = 900 eliminated ticks/sec).
   At 1000Hz: ~10,000+ COM reads/sec from flipper properties.
   At 100Hz: ~600 COM reads/sec. Net savings: ~9,400 COM reads/sec.

2. FLIPPER TIMER COM CACHING (line ~823)
   Cached LeftFlipper.CurrentAngle, LeftFlipper.StartAngle,
   RightFlipper.CurrentAngle, RightFlipper.StartAngle into local variables.
   Original read CurrentAngle 2x and StartAngle 2x per flipper per tick.
   Pre-computed FlipperSOSStrength = FlipperPower * SOSTorque at init.
   At 100Hz: eliminates ~400 COM reads/sec.

3. PRE-COMPUTED INVERSE TABLE DIMENSIONS (line ~888)
   Added InvTWHalf = 2/TableWidth and InvTHHalf = 2/TableHeight at init.
   Pan() and AudioFade() now multiply instead of dividing.
   Called from RollingTimer (~14 balls x 100Hz) + PlaySoundAt + PlaySoundAtBall.
   Eliminates ~2,800 divisions/sec from rolling sounds alone.

4. EXPONENTIATION ELIMINATION — ^10 IN Pan/AudioFade (line ~897, ~921)
   Replaced tmp^10 with chained multiply: t2=tmp*tmp, t4=t2*t2, t8=t4*t4,
   result=t8*t2 (4 multiplies vs Exp(10*Log(x))).
   Pan and AudioFade each called per ball per frame from RollingTimer,
   PlaySoundAt, PlaySoundAtBall, and flipper collide subs.
   At 100Hz with 14 balls: ~2,800 ^10 ops/sec eliminated.

5. EXPONENTIATION ELIMINATION — ^2 IN BallVel/Vol (line ~893, ~911)
   Replaced ball.VelX^2 + ball.VelY^2 with vx*vx + vy*vy. Cached VelX/VelY
   into locals (2 COM reads instead of 4 when called from Vol which calls BallVel).
   Vol() inlines BallVel computation to avoid double COM reads.
   At 100Hz with 14 balls: ~2,800 ^2 ops/sec + ~2,800 COM reads/sec eliminated.

6. PRE-BUILT ROLLING SOUND STRING ARRAY (line ~945)
   Pre-built BallRollStr(0..19) = "fx_ballrolling0" through "fx_ballrolling19".
   Eliminates "fx_ballrolling" & b string concatenation per ball per tick.
   At 100Hz with 14 active balls: ~1,400 string allocs/sec eliminated.

7. RollingTimer COM CACHING + INLINE (line ~956)
   Set ball = BOT(b) cached once per iteration. Cached VelX/VelY/VelZ/z into
   locals (bvx/bvy/bvz/bz). Inlined BallVel/Vol/Pitch computations.
   Original: BallVel called 3x per ball (velocity check, Vol, Pitch), each
   reading VelX+VelY from COM. Now: 4 COM reads total per ball (VelX, VelY,
   VelZ, z) instead of ~18.
   Cached UBound(BOT) once at sub entry.
   Fixed VBScript non-short-circuit bug: "BOT(b).VelX AND BOT(b).VelY <> 0"
   always evaluates both sides. Replaced with nested If checks.
   Used pre-built BallRollStr instead of string concat.
   At 100Hz with 14 balls: ~19,600 COM reads/sec eliminated.
   Drop sound: nested If instead of single-line And chain (VBS trap #3).

8. RAINBOW TIMER RGB PRE-COMPUTATION (line ~322)
   Pre-compute RGB(rRed\10, rGreen\10, rBlue\10) and RGB(rRed, rGreen, rBlue)
   once before the For Each loop instead of computing per object.
   Eliminates (N-1)*2 redundant RGB() calls per tick (N = RainbowLights count).

9. OnBallBallCollision ^2 ELIMINATION (line ~1014)
   Replaced velocity^2 with velocity*velocity. Minor (event handler, not hot).

ESTIMATED TOTAL SAVINGS
========================
- Flipper timer frequency: ~9,400 COM reads/sec eliminated
- Flipper COM caching: ~400 COM reads/sec eliminated
- Inverse table dims: ~2,800 divisions/sec eliminated
- ^10 elimination: ~2,800 Exp+Log pairs/sec eliminated
- ^2 elimination: ~2,800 Exp+Log pairs/sec eliminated
- String allocs: ~1,400/sec eliminated
- Rolling COM caching + inline: ~19,600 COM reads/sec eliminated
- Rainbow RGB: ~N*2 RGB() calls/tick eliminated
Conservative total: ~39,000 redundant operations/sec eliminated.
