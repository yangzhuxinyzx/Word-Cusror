# Subagents

Phase 7 introduces a runtime-level subagent system oriented around document workflows.

## Builtin Profiles

- `doc-explore`
- `workspace-explore`
- `doc-editor`
- `ppt-builder`
- `excel-operator`
- `verification`

Each profile declares:

- `safety`
- `defaultMode`
- `canRunInBackground`
- `prompt`
- `allowedToolIds`

## Runtime Shape

Subagents are managed by [`SubagentManager`](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/agent/subagents/SubagentManager.ts).

Each subagent record includes:

- `subagentId`
- `parentSessionId`
- `parentTurnId`
- `profileId`
- `mode`
- `status`
- `transcriptSessionId`
- `taskId`
- `allowedToolIds`

## Background Support

Background subagents are represented as runtime tasks and also emit task notifications.

Related runtime modules:

- [`TaskRegistry.ts`](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/agent/tasks/TaskRegistry.ts)
- [`TaskRunner.ts`](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/agent/tasks/TaskRunner.ts)
- [`TaskNotifications.ts`](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/agent/tasks/TaskNotifications.ts)

## Transcript Isolation

Each subagent receives an isolated transcript session id and stores its own messages / tool events in:

- [`SessionTranscriptStore.ts`](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/agent/storage/SessionTranscriptStore.ts)

## Current Status

Phase 7 currently provides:

- profile registry
- subagent lifecycle manager
- background task mapping
- transcript isolation
- notification wiring
- AgentSessionEngine APIs for spawn / start / complete / fail / cancel
- ChatPanel command entry for sync and background subagent runs
- isolated child-session execution for `workspace-explore`, `doc-explore`, and `verification`

## Chat Commands

Current command entrypoints:

- `/workspace-explore <request>`
- `/doc-explore <request>`
- `/doc-editor <request>`
- `/ppt-builder <request>`
- `/excel-operator <request>`
- `/verification <request>`
- `/verify <request>`
- `/subagent <profile-id> <request>`
- `/bg-subagent <profile-id> <request>`
- `/bg-workspace-explore <request>`
- `/bg-ppt-builder <request>`

What is not yet built:

- richer profile-specific context adapters beyond the current document/workspace bootstrap
- UI for starting and observing subagents
- automatic result merge-back into the main chat stream
