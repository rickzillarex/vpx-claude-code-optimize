X-Files VBS -- Optimization Changes
===============================================

1. LAMPS CALLBACK COM READ CACHING (line ~586)
   The Lamps sub reads Controller.Lamp(N) for 74 lamps, but many lamps have 2-4
   VPX light objects sharing the same lamp number. Previously, Controller.Lamp(N)
   was called separately for each object (e.g., Controller.Lamp(52) called 3 times
   for L52, L52a, L52b). Now each Controller.Lamp(N) is read once into local var
   's' and assigned to all associated objects.
   Eliminated ~65 redundant Controller.Lamp() COM reads per tick.
   At 100Hz: ~6,500 COM reads/sec eliminated.

2. PAN ^10 ELIMINATION (line ~882)
   Replaced tmp^10 with chained multiplies: t2=tmp*tmp, t4=t2*t2, t8=t4*t4,
   result=t8*t2. Pre-computed InvTWHalf = 2/table1.width at module level.
   Eliminates per-call division and exponentiation.
   At 100Hz with 5 balls: ~500 ^10 ops/sec + ~500 divisions/sec eliminated.

3. VOL ^2 ELIMINATION (line ~878)
   Changed Vol = BallVel(ball)^2/1000 to bv*bv/1000.

4. BALLVEL ^2 ELIMINATION (line ~896)
   Changed ball.VelX^2 + ball.VelY^2 to vx*vx + vy*vy with cached COM reads.

5. ROLLINGTIMER FULL REWRITE (line ~914)
   - Pre-built BallRollStr() array at init. Eliminates "fx_ballrolling" & b
     string concatenation every tick per ball.
   - Set ball = BOT(b) caches ball reference.
   - Cached ball.VelX/VelY/z into locals. Eliminates ~8 COM reads per ball per tick.
   - Inlined BallVel/Vol/Pitch computations. Eliminates 3 function calls + 6
     redundant COM reads per ball.
   - Cached UBound(BOT) once.
   At 100Hz with 5 balls: ~2,500 COM reads/sec + ~500 string allocs/sec +
   ~1,500 function calls/sec eliminated.

6. GATETIMER GATE ANGLE WRITE GUARDS (line ~773)
   Added lastGate3Angle/lastGate4Angle/lastSw25Angle/lastSw27Angle/lastGate6Angle
   tracking. Only writes RotX when gate angle changes. Gates are stationary most of
   gameplay, so this eliminates 5 redundant COM writes per frame.
   At 60fps: ~300 COM writes/sec eliminated.

7. GATETIMER TEXTURE SWAP GUARDS (line ~773)
   Added lastAlienState and lastCabState tracking. Alien baby texture swapping
   (3 .image writes) and file cabinet texture swapping (2 .image writes) now only
   execute when the underlying light states actually change. Previously, 5 light
   state reads + 3-5 .image writes happened every single frame unconditionally.
   Image writes are among the most expensive COM operations (GPU texture swap).
   At 60fps: ~300 redundant .image writes/sec eliminated (alien) +
   ~120 redundant .image writes/sec eliminated (cabinet).

8. ONBALLBALLCOLLISION ^2 ELIMINATION (line ~945)
   Changed velocity^2 to velocity*velocity.


ESTIMATED TOTAL SAVINGS
========================
Conservative total: ~12,000 redundant operations/sec eliminated.
Primary wins: Lamps COM caching (~6,500), RollingTimer inlining (~4,500),
GateTimer texture/gate guards (~720), Pan ^10 elimination (~1,000).
