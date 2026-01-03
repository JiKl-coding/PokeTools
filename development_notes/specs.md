# SPECS for SDD with AI agent (GPT-5.2 at the time of development)

1. `north-star.md`
   - Defines **why the project exists**, its vision, goals, and non-goals.
   - Sets the overall direction and constraints for all other specs.

2. `data-contract.md`
   - Defines **what data exists** in the system.
   - Specifies raw entities, derived models, fields, types, and relationships.
   - Acts as the single source of truth for data shape.

3. `fetch.spec.md`
   - Defines **how raw data is fetched and cached** from PokéAPI.
   - Covers discovery, TTL-based caching, file structure, and CLI behavior.
   - No transformations or exports allowed.

4. `transform.spec.md`
   - Defines **how raw cached data is transformed** into normalized domain models.
   - Covers deterministic rules, learnsets, evolutions, type chart, and metadata.
   - No network access, no Excel logic.

5. `export.spec.md`
   - Defines **how derived models are exported** into `pokedata.xlsx`.
   - Specifies sheets, columns, ordering, and allowed formatting.
   - No fetching or business logic.

**Prompt example:**
Implement 1. MVP Fetch – Initial Test according to #1_fetch_prompt.md. FOLLOW Requirements & Source of Truth defined in the #1_fetch_prompt.md. Do not implement any other phases or additional logic.