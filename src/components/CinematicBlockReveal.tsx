import { useEffect, useMemo, useRef } from 'react'
import { motion, Variants } from 'framer-motion'

interface CinematicBlockRevealProps {
  /**
   * 原始内容（支持 Markdown 或 HTML）
   */
  content: string
  /**
   * 是否将内容解析为 HTML 块。若为 false，则按空行拆分段落。
   */
  parseAsHtml?: boolean
  /**
   * 每段之间的阶梯延迟，单位秒。
   */
  stagger?: number
  className?: string
  onComplete?: () => void
  /**
   * 最大动画段落数量，避免一次性内容过多导致性能问题
   */
  maxBlocks?: number
}

const BLOCK_NODE_SELECTOR = 'p, h1, h2, h3, h4, h5, h6, li, blockquote, table, ul, ol, pre'

const splitTextBlocks = (text: string) => {
  return text
    .split(/\n\s*\n/)
    .map(block => block.trim())
    .filter(Boolean)
}

const extractHtmlBlocks = (html: string) => {
  const parser = new DOMParser()
  const doc = parser.parseFromString(html, 'text/html')
  const nodes = Array.from(doc.body.querySelectorAll<HTMLElement>(BLOCK_NODE_SELECTOR))
  if (!nodes.length) {
    const fallback = doc.body.innerHTML.trim()
    return fallback ? [fallback] : []
  }
  return nodes.map(node => node.outerHTML.trim()).filter(Boolean)
}

const createVariants = (stagger: number): Variants => ({
  hidden: {
    opacity: 0,
    y: 36,
    filter: 'blur(12px)',
  },
  visible: (index: number = 0) => ({
    opacity: 1,
    y: 0,
    filter: 'blur(0px)',
    transition: {
      delay: index * stagger,
      type: 'spring',
      stiffness: 110,
      damping: 16,
      mass: 0.9,
    },
  }),
  exit: {
    opacity: 0,
    y: -12,
    filter: 'blur(6px)',
    transition: { duration: 0.2, ease: 'easeInOut' },
  },
})

const CinematicBlockReveal: React.FC<CinematicBlockRevealProps> = ({
  content,
  parseAsHtml = false,
  stagger = 0.12,
  className,
  onComplete,
  maxBlocks = 80,
}) => {
  const hasCompletedRef = useRef(false)
  const variants = useMemo(() => createVariants(stagger), [stagger])

  const blocks = useMemo(() => {
    if (!content) return []
    const rawBlocks = parseAsHtml ? extractHtmlBlocks(content) : splitTextBlocks(content)
    return rawBlocks.slice(0, maxBlocks)
  }, [content, parseAsHtml, maxBlocks])

  useEffect(() => {
    hasCompletedRef.current = false
  }, [blocks.length])

  if (!blocks.length) return null

  return (
    <div className={`pointer-events-none select-none ${className ?? ''}`}>
      {blocks.map((block, index) => (
        <motion.div
          key={`${index}-${block.length}`}
          custom={index}
          variants={variants}
          initial="hidden"
          animate="visible"
          exit="exit"
          style={{ willChange: 'transform, opacity, filter' }}
          onAnimationComplete={() => {
            if (index === blocks.length - 1 && !hasCompletedRef.current) {
              hasCompletedRef.current = true
              onComplete?.()
            }
          }}
          className="mb-3 last:mb-0"
        >
          <div
            className="leading-[1.7] text-zinc-800 prose prose-slate max-w-none"
            dangerouslySetInnerHTML={{ __html: block }}
          />
        </motion.div>
      ))}
    </div>
  )
}

export default CinematicBlockReveal

