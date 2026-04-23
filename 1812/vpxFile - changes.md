1812 vpxFile VBS — Optimization Changes
===============================================

NOTE: This table was already partially optimized (OPT-1 through OPT-6 in the
original). The changes below are incremental improvements on top of those.

1. DICTIONARY ELIMINATION IN AddLamp (line ~587)
   Replaced AddLamp(nr, object) with AddLampT(nr, object, objType) that takes
   the type code (1=Light, 2=Flasher, 3=Primitive) as a direct parameter.
   Eliminated the Scripting.Dictionary (LampObjTypeDict) and its per-call
   string key construction (CStr(nr) & "_" & ObjName(object)) plus .Item()
   hash lookup. Registration code (RegisterLampObj) removed entirely.
   At 33Hz: 35 calls/tick x 33 ticks/sec = ~1,155 hash lookups + ~1,155
   string concatenations eliminated per second.

2. HideLamp DELTA GUARD (line ~647)
   Added lastHideLamp sentinel variable. HideLamp now exits immediately if
   the enabled state hasn't changed. Eliminates 3 redundant .visible COM
   writes per tick when state is stable (the common case).
   At 33Hz: ~198 COM writes/sec eliminated when stable.

3. AddLampT FlashLevel BOUNDARY GUARD (line ~655)
   Added boundary checks: skip FlashLevel arithmetic when already at 0 (off)
   or 1 (on). When lamps are fully on or fully off (majority of gameplay),
   eliminates 1 subtraction + 1 comparison per call.
   At 33Hz: up to ~1,155 arithmetic ops/sec eliminated during steady state.

4. EXPONENTIATION ELIMINATION IN COLLISION HANDLERS (lines ~812, ~825, ~837)
   OnBallBallCollision: velocity^2 replaced with velocity*velocity.
   RubberPosts_Hit/Rubbers_Hit: activeball.velx^2 + vely^2 replaced with
   cached locals rpvx*rpvx + rpvy*rpvy. Also caches VelX/VelY into locals
   (2 COM reads instead of 2 per expression).
   These are event handlers (not hot), but the fix is free and correct.

ESTIMATED TOTAL SAVINGS
========================
- Dictionary + string ops: ~2,310/sec eliminated (hash lookups + string concat)
- Redundant COM writes: ~198/sec eliminated (HideLamp guard)
- FlashLevel arithmetic: ~1,155/sec eliminated (boundary guard, steady state)
- Exponentiation: minor (event handlers, ~5-15/sec)
Conservative total: ~3,600 redundant operations/sec eliminated.
