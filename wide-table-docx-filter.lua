function setTableColspecs(el)
    -- Determine the number of columns
    local num_columns = #el.colspecs

    -- Set equal width for each column (relative width)
    local custom_width = 1.0 / num_columns -- Distribute equally

    -- Apply custom widths and alignment (e.g., centered alignment)
    for i = 1, num_columns do
        el.colspecs[i] = {pandoc.AlignLeft, el.colspecs[i].width}
    end

    return el
end

function calculateSpans(charCount, threshold)
    -- Determines the span based on character count
    -- Increase `threshold` if you want larger content to span more cells
    local span = math.ceil(charCount / threshold)
    return math.max(1, span) -- Ensure at least a span of 1
end

function adjustCellSpans(cell, threshold)
    local cellContent = pandoc.utils.stringify(cell.contents)
    local charCount = #cellContent

    -- Set row and column span based on character count
    cell.row_span = calculateSpans(charCount, threshold.row)
    cell.col_span = calculateSpans(charCount, threshold.col)

    return cell
end

-- Function to escape special XML characters
function escapeXML(str)
    if type(str) == "string" then
        str = str:gsub("&", "&amp;")
        str = str:gsub("<", "&lt;")
        str = str:gsub(">", "&gt;")
        str = str:gsub('"', "&quot;")
        str = str:gsub("'", "&apos;")
    end
    return str
end

function wrapCellContentWithStyles(cell, style, isHeader)
    if cell and cell.contents then
        local contents = cell.contents

        -- Initialize styledContent as an empty table to hold inline elements
        local styledContent = {}

        -- Walk through each block within the cell's contents
        for _, block in ipairs(contents) do
            if block.t == "Para" or block.t == "Plain" then
                -- Handle block content by wrapping inline elements within it
                for _, inline in ipairs(block.content) do
                    table.insert(styledContent, pandoc.Span(inline, {
                        ['custom-style'] = style
                    }))
                end
            else
                -- If the block is not a paragraph or plain, treat it as inline (less common case)
                table.insert(styledContent, pandoc.Span(pandoc.utils.stringify(block), {
                    ['custom-style'] = style
                }))
            end
        end

        -- Update the cell's contents with the newly styled inline elements
        cell.contents = {pandoc.Plain(styledContent)}
    end
    return cell
end

function processCells(cells, style, isHeader)
    local threshold = {
        row = 10,
        col = 15
    } -- Adjust thresholds as needed
    for _, cell in ipairs(cells) do
        -- Process each block in the cell
        if cell then
            wrapCellContentWithStyles(cell, style, isHeader)
            -- adjustCellSpans(cell, threshold)
        end
    end
    return cells
end

function processRows(rows, style, isHeader)
    for _, row in ipairs(rows) do
        if row then -- Ensure the row is valid
            processCells(row.cells, style, isHeader) -- Process the cells in each row
        end
    end
    return rows
end

function replaceCaption(container, find, replace)
  for i, elem in ipairs(container) do
    if elem.t == "Str" then
      elem.text = elem.text:gsub(find, replace)
    elseif elem.c then
      replaceCaption(elem.c, find, replace)
    end
  end
end

function Table(el)
    if FORMAT:match("docx") then
        local caption_text = pandoc.utils.stringify(el.caption.long)
        local style = caption_text:match("{#cellstyle:(.-)}")

        if style then
            replaceCaption(el.caption.long, "%s*{#cellstyle:.-}", "")
            setTableColspecs(el) -- to set the table column width

            -- Process the table head
            if el.head and el.head.rows and type(el.head.rows) == "table" then
                processRows(el.head.rows, style, true)
            end

            -- Process the table body
            for _, body in ipairs(el.bodies) do
                if body and body.body and type(body.body) == "table" then
                    processRows(body.body, style, false)
                end
            end

            -- Process the table foot (if any)
            if el.foot and el.foot.rows and type(el.foot.rows) == "table" then
                -- processRows(el.foot.rows, style, false)
            end
        end
    end
    return el
end

function Div(el)
    if FORMAT:match("docx") and el.classes:includes("landscape") then
        table.insert(el.content, 1, pandoc.RawBlock("openxml",
            '<w:p><w:pPr><w:sectPr><w:pgSz w:h="15840" w:w="12220" w:orient="portrait"/><w:pgMar w:top="1400" w:right="1800" w:bottom="1380" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:pPr></w:p>'))
        table.insert(el.content, pandoc.RawBlock("openxml",
            '<w:p><w:pPr><w:sectPr><w:pgSz w:w="15840" w:h="12220" w:orient="landscape"/><w:pgMar w:top="1440" w:right="720" w:bottom="1440" w:left="720" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:pPr></w:p>'))
    end
    return el
end
