export interface ReactiveCompactSignal {
  shouldCompact: boolean
  reason?: string
}

export function shouldReactiveCompact(params: {
  wasTruncated?: boolean
  modelRejected?: boolean
  promptTooLarge?: boolean
}): ReactiveCompactSignal {
  if (params.promptTooLarge) {
    return { shouldCompact: true, reason: 'prompt_too_large' }
  }
  if (params.modelRejected) {
    return { shouldCompact: true, reason: 'model_rejected' }
  }
  if (params.wasTruncated) {
    return { shouldCompact: true, reason: 'response_truncated' }
  }
  return { shouldCompact: false }
}
