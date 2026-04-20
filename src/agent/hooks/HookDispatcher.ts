import type { RegisteredRuntimeHook, RuntimeHookEvent } from './HookRegistry'

export class HookDispatcher {
  async dispatch(
    event: RuntimeHookEvent,
    hooks: RegisteredRuntimeHook[],
    invoke: (hook: RegisteredRuntimeHook) => Promise<void>,
  ): Promise<void> {
    const matched = hooks.filter((hook) => hook.event === event)
    for (const hook of matched) {
      await invoke(hook)
    }
  }
}
