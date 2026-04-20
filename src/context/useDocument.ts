import { useContext } from 'react'
import { DocumentContext } from './DocumentContext'

export function useDocument() {
  const context = useContext(DocumentContext)
  if (!context) {
    throw new Error('useDocument must be used within a DocumentProvider')
  }
  return context
}
