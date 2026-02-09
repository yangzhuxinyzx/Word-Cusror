-- pagebreak.lua
-- 将 <!-- pagebreak -->, \newpage 和 \pagebreak 转换为 Word 分页符

function RawBlock(el)
    -- 处理 <!-- pagebreak -->
    if el.text == '<!-- pagebreak -->' then
        return pandoc.RawBlock('openxml', '<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
    end
    
    -- 处理 \newpage 和 \pagebreak
    if el.text == '\\newpage' or el.text == '\\pagebreak' then
        return pandoc.RawBlock('openxml', '<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
    end
end

function Div(el)
    -- 处理 ::: pagebreak 这样的 div 语法
    if el.attr and el.attr.classes:includes('pagebreak') then
        return pandoc.RawBlock('openxml', '<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
    end
end

local stringify = pandoc.utils.stringify

function Pandoc(doc)
    -- 查找文档中特定的标记位置
    local blocks = {}
    local toc_inserted = false
    local toc_marker = "{{TOC}}"

    for i, el in pairs(doc.blocks) do
        if el.t == "Para" then
        local content = stringify(el)
        if content == toc_marker then
            table.insert(blocks, pandoc.RawBlock('openxml', 
            [[<w:sdt>
                <w:sdtPr>
                <w:docPartObj>
                    <w:docPartGallery w:val="Table of Contents"/>
                    <w:docPartUnique/>
                </w:docPartObj>
                </w:sdtPr>
                <w:sdtContent>
                <w:p>
                    <w:pPr>
                    <w:pStyle w:val="TOC"/>
                    </w:pPr>
                    <w:r>
                    <w:rPr>
                        <w:rFonts w:hint="eastAsia"/>
                    </w:rPr>
                    <w:t xml:space="preserve">目录</w:t>
                    </w:r>
                </w:p>
                <w:p>
                    <w:r>
                    <w:fldChar w:fldCharType="begin" w:dirty="true"/>
                    <w:instrText xml:space="preserve">TOC \o "1-4" \h \z \u</w:instrText>
                    <w:fldChar w:fldCharType="separate"/>
                    <w:fldChar w:fldCharType="end"/>
                    </w:r>
                </w:p>
                </w:sdtContent>
            </w:sdt>]]))
            toc_inserted = true
        else
            table.insert(blocks, el)
        end
        else
        table.insert(blocks, el)
        end
    end

    return pandoc.Pandoc(blocks, doc.meta)
end

function CodeBlock(block)
    if block.text:match '^!include ' then
      local filename = block.text:match '^!include (%S+)'
      return pandoc.read(io.open(filename):read('*a'), 'markdown').blocks
    end
  end
  