# Skills instaladas

Copiadas desde GitHub (rama por defecto, commit indicado):

- https://github.com/analystacademy/claude-skills @ 736213a
- https://github.com/blader/humanizer @ 225a6f3
- https://github.com/bradautomates/claude-video @ 03ceb42
- https://github.com/mvanhorn/last30days-skill @ 50998cd
- https://github.com/pbakaus/impeccable @ abf5779 (actualizado)
- https://github.com/remotion-dev/skills @ 7ba4dcf
- https://github.com/zarazhangrui/frontend-slides @ 9906a34

Notas:
- remotion-dev/remotion: sus skills oficiales viven en remotion-dev/skills (`npx skills add remotion-dev/skills`).
- last30days: se omitió la carpeta assets/ (solo imágenes de demo).
- impeccable: no se instalaron sus hooks (SessionStart/PostToolUse/Stop) de .claude/settings.json; ver https://github.com/pbakaus/impeccable si los quieres.

## Segunda tanda
- https://github.com/leonxlnx/taste-skill @ f624f37: 14 skills, carpetas nombradas con su `name` (como hace `npx skills add`).
- https://github.com/nextlevelbuilder/ui-ux-pro-max-skill @ 56ea9cf: 7 skills de su `.claude/skills/` (lo mismo que instala `uipro init --ai claude`).
- https://github.com/VoltAgent/awesome-claude-code-subagents @ 408d3c9: 163 subagentes de las 10 categorías en `.claude/agents/` (incluye agent-installer).
- https://github.com/headroomlabs-ai/headroom: NO instalado aquí; es un proxy local (`pip install "headroom-ai[all]"` + `headroom wrap claude`), pensado para tu PC.
- https://github.com/nidhinjs/prompt-master @ 769bf64: skill `prompt-master` (SKILL.md + references).
