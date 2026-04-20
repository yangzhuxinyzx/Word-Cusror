# Workspace Skills

Word-Cursor now supports workspace-defined skills in addition to builtin skills.

## Supported Locations

Place skill manifests in either of these paths under the active workspace:

- `.word-cursor/skills.json`
- `.word-cursor/skills/*.json`

## Supported Shape

Each file may contain:

- a single skill object
- an array of skill objects
- an object with a `skills` array

## Minimal Skill Example

```json
{
  "id": "policy-proofread",
  "displayName": "Policy Proofread",
  "description": "Proofread policy documents and normalize official wording.",
  "executionKind": "prompt_transform",
  "safety": "verification",
  "prompt": "Proofread the current policy document, preserve meaning, normalize official language, and flag ambiguous wording.",
  "toolIds": ["word.read_selection", "word.read_outline", "review", "replace"],
  "invocation": {
    "slashCommands": ["policy-proofread", "政策校对"]
  }
}
```

## Notes

- Workspace skills are currently normalized to `prompt_transform`.
- Skill `id` must match `^[a-z0-9][a-z0-9-_]{1,63}$`.
- If a workspace skill has no `prompt`, it will still load, but a warning will appear in `/skills`.
- If a skill declares `toolIds`, ChatPanel will block tools outside that allowlist for that turn.

## Discovery

Use these commands in chat:

- `/skills`
- `/init`

See [`workspace-skills.example.json`](/C:/Users/yangz/Desktop/Github/Word-Cursor/docs/workspace-skills.example.json) for a fuller example.
