export type RuntimeHookEvent =
  | 'PreToolUse'
  | 'PostToolUse'
  | 'PostToolUseFailure'
  | 'PermissionDenied'
  | 'InstructionsLoaded'

export interface RegisteredRuntimeHook {
  id: string
  event: RuntimeHookEvent
  description: string
}

export class HookRegistry {
  private hooks = new Map<string, RegisteredRuntimeHook>()

  register(hook: RegisteredRuntimeHook): void {
    this.hooks.set(hook.id, { ...hook })
  }

  list(): RegisteredRuntimeHook[] {
    return Array.from(this.hooks.values())
  }
}

