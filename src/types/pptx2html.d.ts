declare module 'dimple' {
  const dimple: any
  export default dimple
}

declare module 'pptx2html' {
  const renderPptx: (
    pptx: ArrayBuffer,
    resultElement: Element | string,
    thumbElement?: Element | string
  ) => Promise<number>
  export default renderPptx
}



