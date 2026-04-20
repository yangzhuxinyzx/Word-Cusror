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
  baseSpeed = 3,   // 提高默认速度：每帧3个字符
  maxSpeed = 12,   // 提高最大速度：积压时每帧12个字符
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
      // 根据积压量动态调整速度，更积极地消费队列
      if (backlog > 100) charsToTake = maxSpeed
      else if (backlog > 50) charsToTake = Math.min(maxSpeed, baseSpeed + 6)
      else if (backlog > 20) charsToTake = Math.min(maxSpeed, baseSpeed + 3)

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
      // 如果 text 变空，但队列中还有内容，继续处理队列
      if (queueRef.current.length > 0 && processingRef.current) {
        prevLengthRef.current = 0
        return // 不重置，让队列继续处理
      }
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
      {stableText}
      {animatedChars.map(({ char, id }) => (
        <span key={id} className="cinematic-char">
          {char}
        </span>
      ))}
    </div>
  )
}

