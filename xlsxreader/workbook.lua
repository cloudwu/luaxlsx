local zip = require "zip"
local xml = require "xlsxreader.xml"
local print_r = require "xlsxreader.print_r"

local M = {}

local function read_xml(self, name)
	local data = self.archive:read(name)
	return xml.collect(data)
end

local function read_sharedstrings(sharedstrings, result)
	assert(sharedstrings and sharedstrings.xml == "sst" , "Invalid sharedstrings.xml")
	for _, v in ipairs(sharedstrings) do
		if v.xml == "si" and v[1] then
			if v[1].xml == "t" then
				table.insert(result, v[1][1] or "")
			else
				-- NOTE: remove font <r> <rPr> ...
				-- http://msdn.microsoft.com/en-us/library/office/gg278314(v=office.15).aspx
				local tmp = {}
				for _, v in ipairs(v) do
					if v.xml == "r" then
						for _, v in ipairs(v) do
							if v.xml == "t" then
								table.insert(tmp, v[1])
							end
						end
					end
				end
				table.insert(result, table.concat(tmp))
			end
		end
	end
end

local STYLES = {}

do
	local function define(tags, tag, style)
		STYLES[tags] = function (xmltbl, result)
			for _,v in ipairs(xmltbl) do
				if v.xml == tag then
					table.insert(result, xml.extract(v, style))
				end
			end
		end
	end

	define("fonts", "font", { color = true, name = "val", sz = "val", charset = "val", family = "val", charset = "val" })

	function STYLES.fills(xmltbl, result)
		for _,v in ipairs(xmltbl) do
			if v.xml == "fill" then
				for _,v in ipairs(v) do
					if v.xml == "patternFill" then
						local tmp = { patternType = v.patternType }
						for _,v in ipairs(v) do
							tmp[v.xml] = v
						end
						table.insert(result, tmp)
					end
				end
			end
		end
	end
	-- todo: support ignore style
	-- ignore numFmts
	-- ignore borders
	-- ignore dxfs
	-- ignore cellStyles
	-- ignore cellStyleXfs
	function STYLES.cellXfs(xmltbl, result)
		for _,v in ipairs(xmltbl) do
			if v.xml == "xf" then
				table.insert(result, v)
			end
		end
	end
end

local function parser(tbl, result, method)
	for _, v in ipairs(tbl) do
		local f = method[v.xml]
		if f then
			local r = result[v.xml] or {}
			result[v.xml] = r
			f(v, r)
		end
	end
end

-- http://www.officeopenxml.com/SSstyles.php
local function read_styles(styles, result)
	assert(styles and styles.xml == "styleSheet" , "Invalid sharedstrings.xml")
	parser(styles, result, STYLES)
end

local function gen_styles(styles, result)
	for _,v in ipairs(styles.cellXfs) do
		local s = {}
		local font = tonumber(v.fontId)
		local fill = tonumber(v.fillId)
		s.font = styles.fonts[font and font + 1]
		s.fill = styles.fills[fill and fill + 1]
		for _,v in ipairs(v) do
			if v.xml == "alignment" then
				s.alignment = v
			end
		end
		table.insert(result, s)
	end
end

local SHEET = {}

do
	function SHEET.sheetData(xmltbl, result)
		local data = {}
		local rows = {}
		for _, v in ipairs(xmltbl) do
			if v.xml == "row" then
				if v.customHeight and v.ht then
					rows[tonumber(v.r)] = tonumber(v.ht)
				end
				for _,v in ipairs(v) do
					if v.xml == "c" then
						local value = { r = v.r, s = v.s , t = v.t }
						table.insert(data, value)
						for _, v in ipairs(v) do
							value[v.xml] = v[1]
						end
					end
				end
			end
		end
		result.data = data
		result.rows = rows
	end

	function SHEET.dimension(xmltbl, result)
		result.ref = xmltbl.ref
	end

	function SHEET.cols(xmltbl, result)
		for _,v in ipairs(xmltbl) do
			if v.xml == "col" then
				local min = tonumber(v.min)
				local max = tonumber(v.max)
				for i=min, max do
					result[i] = tonumber(v.width)
				end
			end
		end
	end
end

local function load_sheet(self, result)
	local tbl = read_xml(self, "xl/" .. result.filename)
	result.filename = nil
	tbl = tbl[2]
	assert(tbl.xml == "worksheet")

	parser(tbl, result, SHEET)
	result.rows, result.sheetData.rows = result.sheetData.rows
	result.sheetData, result.sheetData.data = result.sheetData.data
	result.dimension, result.dimension.ref = result.dimension.ref

	for _, v in ipairs(result.sheetData) do
		local s = tonumber(v.s)
		if s then
			v.s = s
		end
		if v.t == "s" then
			v.v = self.sharedstrings[tonumber(v.v)+1]
		end
	end
end

function M.load(filename)
	local self = {}
	self.archive = assert(zip.unzip(filename), "Can't open " ..  filename)
	self.sharedstrings = {}
	local ok, sharedstrings = pcall(read_xml, self, "xl/sharedstrings.xml")
	if ok then
		read_sharedstrings(sharedstrings[2], self.sharedstrings)
	end

	self.styles = {}
	local ok, styles = pcall(read_xml, self, "xl/styles.xml")
	if ok then
		local tmp = {}
		read_styles(styles[2], tmp)
		gen_styles(tmp, self.styles)
	end

	local sheets = {}
	do
		local workbook = read_xml(self, "xl/workbook.xml")
		workbook = workbook[2]
		assert(workbook.xml == "workbook")
		local workbook_rels = read_xml(self, "xl/_rels/workbook.xml.rels")
		local rels = {}
		assert (workbook_rels[2].xml == "Relationships")
		for _,v in ipairs(workbook_rels[2]) do
			rels[v.Id] = v.Target
		end
		for _,v in ipairs(workbook) do
			-- only support sheets
			if v.xml == "sheets" then
				for _,v in ipairs(v) do
					if v.xml == "sheet" then
						sheets[tonumber(v.sheetId)] = {
							name = v.name,
							filename = rels[v["r:id"]]
						}
					end
				end
			end
		end
		for _,v in pairs(sheets) do
			load_sheet(self, v)
		end
	end

	self.archive:close()

	return sheets, self.styles
end

return M
