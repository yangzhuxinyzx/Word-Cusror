import { useCallback, useEffect, useMemo, useRef, useState } from 'react'

interface CinematicTyperProps {
  text: string
  isStreaming: boolean
  baseSpeed?: number
  maxSpeed?: number
}

interface AnimatedChar {
  char: string
  id: number
}

const MAX_ANIMATED_CHARS = 20

export default function CinematicTyper({
  text,
  isStreaming,
  baseSpeed = 1,
  maxSpeed = 4,
}: CinematicTyperProps) {
  const [displayed, setDisplayed] = useState('')
  const [animatedChars, setAnimatedChars] = useState<AnimatedChar[]>([])

  const queueRef = useRef<string[]>([])
  const prevLengthRef = useRef(0)
  const frameRef = useRef<number>()
  const processingRef = useRef(false)
  const charIdRef = useRef(0)

  const stopLoop = () => {
    if (frameRef.current) {
      cancelAnimationFrame(frameRef.current)
      frameRef.current = undefined
    }
  }

  const processChunk = useCallback(
    (timestamp: number) => {
      const queue = queueRef.current
      if (queue.length === 0) {
        processingRef.current = false
        stopLoop()
        return
      }

      let charsToTake = baseSpeed
      const backlog = queue.length
      if (backlog > 80) charsToTake = maxSpeed
      else if (backlog > 50) charsToTake = Math.min(maxSpeed, baseSpeed + 2)
      else if (backlog > 30) charsToTake = Math.min(maxSpeed, baseSpeed + 1)

      const chunk = queue.splice(0, charsToTake)

      setDisplayed(prev => prev + chunk.join(''))
      setAnimatedChars(prev => {
        const newEntries = chunk.map(char => ({ char, id: charIdRef.current++ }))
        const merged = [...prev, ...newEntries]
        return merged.slice(-MAX_ANIMATED_CHARS)
      })

      frameRef.current = requestAnimationFrame(processChunk)
    },
    [baseSpeed, maxSpeed]
  )

  useEffect(() => {
    if (!text) {
      setDisplayed('')
      setAnimatedChars([])
      queueRef.current = []
      prevLengthRef.current = 0
      processingRef.current = false
      stopLoop()
      return
    }

    const prevLength = prevLengthRef.current

    if (text.length < prevLength) {
      // 新会话或内容被重置
      setDisplayed(text)
      setAnimatedChars(prev => prev.slice(-MAX_ANIMATED_CHARS))
      queueRef.current = []
      prevLengthRef.current = text.length
      processingRef.current = false
      stopLoop()
      return
    }

    if (text.length > prevLength) {
      const newChars = Array.from(text.slice(prevLength))
      queueRef.current.push(...newChars)
      prevLengthRef.current = text.length

      if (!processingRef.current) {
        processingRef.current = true
        frameRef.current = requestAnimationFrame(processChunk)
      }
    }
  }, [text, processChunk])

  useEffect(() => {
    return () => {
      stopLoop()
    }
  }, [])

  const stableText = useMemo(() => {
    const animatedLength = animatedChars.length
    if (animatedLength === 0) return displayed
    return displayed.slice(0, Math.max(0, displayed.length - animatedLength))
  }, [displayed, animatedChars])

  const showThinking = isStreaming && displayed.length === 0 && animatedChars.length === 0 && queueRef.current.length === 0

  return (
    <div className="cinematic-typer">
      {showThinking && (
        <div className="cinematic-thinking">
          <span />
          <span />
          <span />
          <span className="cinematic-thinking-text">AI 正在思考…</span>
        </div>
      )}
      {stableText}
      {animatedChars.map(({ char, id }) => (
        <span key={id} className="cinematic-char">
          {char}
        </span>
      ))}
    </div>
  )
}

