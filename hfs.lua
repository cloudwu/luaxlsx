--[[
	hfs format like this:
MFS2
[!sheet]
name:xxx
[data]
A1:xxx
A2:xxx
..
[fill]
....1111....
[layout]
col:1:xxx
col:2:xxx
row:1:xxx
row:2:xxx
[fvalue]
A1:xxx
A2:xxx
[!styles]

In data sector, xxx can be number, string and formula ( leading by = )
In fill sector, A bitmap ref styles.
The excape char is \ (C rules)

]]

local rd = require "xlsxreader.workbook"
local wt = require "xlsxwriter.workbook"
local xml = require "xlsxreader.xml"
local util = require "xlsxwriter.Utility"
local base64 = require "base64"

-------------- xlsx to hfs --------------------

local escape_value_tbl = { ["\\"] = "\\\\" , ["\n"] = "\\n", ["\r"] = "\\r" }
local function escape_value(s)
	s = xml.unescape(s or "")
	s = s:gsub("[\\\r\n]", escape_value_tbl)
	s = s:gsub("^=", "\\=")
	return s
end

local fillc = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ~!@#$%^&*-=+:"
local function fill(sheet, dimension)
	local row, col = util.cell_to_rowcol(dimension:match(":(%w+)$"))
	local tmp = {}
	for i=0,row do
		local line = {}
		table.insert(tmp, line)
		for j = 0, col do
			line[j+1] = "."
		end
	end
	for _,v in ipairs(sheet.sheetData) do
		if v.s then
			local c
			if v.s > #fillc then
				c = "?"
			else
				c = fillc:sub(v.s+1,v.s+1)
			end

			local row, col = util.cell_to_rowcol(v.r)
			tmp[row+1][col+1] = c
		end
	end
	for i=0,row do
		tmp[i+1] = table.concat(tmp[i+1])
	end
	return table.concat(tmp,"\n")
end

local function sort_lines(prefix,rows)
	if rows == nil then
		return ""
	end
	local row = {}
	for k,v in pairs(rows) do
		table.insert(row, { k, v })
	end
	table.sort(row, function(a,b) return a[1] < b[1] end)
	local formatstr = prefix .. "%d:%g\n"
	for k,v in ipairs(row) do
		row[k] = string.format(formatstr,v[1],v[2])
	end

	return table.concat(row)
end

local function dump_table(t)
	local tmp = {}
	for k,v in pairs(t) do
		if k ~= "xml" then
			local value
			local format
			if type(v) == "table" then
				format = "%s={ %s }"
				v = dump_table(v)
			else
				if tonumber(v) then
					format = "%s=%g"
				else
					format = '%s="%s"'
				end
			end
			table.insert(tmp, string.format(format,k,v))
		else
			t.xml = nil
		end
	end
	return table.concat(tmp, ",")
end

local function x2h(xlsxname, hfsname)
	local sheets, styles = rd.load(xlsxname)
	local f = assert(io.open(hfsname, "wb"))
	f:write "MFS2\n"
	for _,v in pairs(sheets) do
		-- header
		f:write(string.format("[!sheet]\nname:%s\n[data]\n", v.name))
		-- data
		local tmp = {}
		for _,v in ipairs(v.sheetData) do
			if v.f then
				table.insert(tmp, string.format("%s:=%s\n", v.r, v.f))
			else
				local s = escape_value(v.v)
				table.insert(tmp, string.format("%s:%s\n", v.r, s))
			end
		end
		f:write(table.concat(tmp))
		--fill
		f:write "[fill]\n"
		f:write(fill(v, v.dimension))
		--layout
		f:write "\n[layout]\n"
		f:write(sort_lines("r", v.rows))
		f:write(sort_lines("c", v.cols))
		--fvalue (function calc value)
		f:write "[fvalue]\n"
		local tmp = {}
		for _,v in ipairs(v.sheetData) do
			if v.f then
				local s = escape_value(v.v)
				table.insert(tmp, string.format("%s:%s\n", v.r, s))
			end
		end
		f:write(table.concat(tmp))
	end
	--styles
	f:write "[!styles]\n"
	for _,v in ipairs(styles) do
		f:write(dump_table(v))
		f:write("\n")
	end
	f:close()
end

----------------- hfs to xlsx -------------------

local function hfs_parser(hfsname)
	local f = assert(io.open(hfsname, "rb"), "can't open "..hfsname)
	local result = {}
	local stack
	local tag = f:read "*l"
	assert(tag == "MFS2", "not a hfs file")
	for line in f:lines() do
		local name = line:match("^%[(!?%w+)%]")
		if name then
			if name == "!sheet" then
				local current = {}
				table.insert(result, current)
				stack = { current , true }
			elseif name == "!styles" then
				local current = {}
				result.styles = current
				stack = { current , true }
			else
				assert(stack, "Need [!sheet] before")
				table.remove(stack)	-- pop top
				local seg = {}
				local top = stack[#stack]
				top[name] = seg
				table.insert(stack, seg)
			end
		elseif stack then
			local top = stack[#stack]
			if top then
				if top == true then
					top = stack[#stack-1]
				end
				table.insert(top, line)
			end
		end
	end

	return result
end

local print_r = require "xlsxreader.print_r"

local unescape_value_tbl = setmetatable({ ["\\n"] = "\n", ["\\r"] = "\r" }, {__index = function(_,k) return k:sub(2) end})
local function unescape_value(s)
	s = s:gsub("(\\.)", unescape_value_tbl)
	return s
end

local patterns = {
	"none",
	"solid",
	"mediumGray",
	"darkGray",
	"lightGray",
	"darkHorizontal",
	"darkVertical",
	"darkDown",
	"darkUp",
	"darkGrid",
	"darkTrellis",
	"lightHorizontal",
	"lightVertical",
	"lightDown",
	"lightUp",
	"lightGrid",
	"lightTrellis",
	"gray125",
	"gray0625",
}
for k,v in ipairs(patterns) do
	patterns[v] = k - 1
end

local function new_format(fm, t)
	local alignment = t.alignment
	if alignment then
		fm:set_align_array(alignment)
	end
	local fill = t.fill
	if fill then
		local index = patterns[fill.patternType]
		if index then
			fm:set_pattern(index)
		end
		if index == 1 then
			-- xlsxwriter exchange bg/fg
			fm:set_fg_color(fill.bgColor)
			fm:set_bg_color(fill.fgColor)
		else
			if fill.fgColor then
				fm:set_fg_color(fill.fgColor)
			end
			if fill.bgColor then
				fm:set_bg_color(fill.bgColor)
			end
		end
	end
	local font = t.font
	if font then
		if font.color then
			fm.set_font_color(font.color)
		end
		if font.sz then
			fm.set_font_size(font.sz)
		end
		if font.name then
			fm.set_font_name(font.name)
		end
		if font.charset then
			fm.set_font_charset(font.charset)
		end
	end
	return fm
end

--local fillc = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ~!@#$%^&*-=+:"
local fill_index = {}
for i=1,#fillc do
	fill_index[fillc:sub(i,i)] = i
end
local function convert_fill(str)
	local r = {}
	for i=1,#str do
		r[i] = fill_index[str:sub(i,i)]
	end
	return r
end

local function h2x(hfsname, xlsxname)
	local r = hfs_parser(hfsname)
	-- create sheets
	local wb = wt:new(xlsxname)

	-- load styles
	local styles = {}
	for _,v in ipairs(r.styles) do
		local f = assert(load(string.format("return {%s}", v)),"t")
		local _, t = assert(pcall(f))
		local fm = wb:add_format()
		table.insert(styles, new_format(fm, t))
	end
	for _,sheet in ipairs(r) do
		-- gen fvalue table
		local fvalue = {}
		if sheet.fvalue then
			for _,line in ipairs(sheet.fvalue) do
				local k,v = line:match("^(%w+):(.*)")
				fvalue[k]=v
			end
		end
		-- gen fill table
		local fill = {}
		if sheet.fill then
			for _,line in ipairs(sheet.fill) do
				table.insert(fill, convert_fill(line))
			end
		end

		local arg = {}
		for _,line in ipairs(sheet) do
			local k,v = line:match("^(%w+):(.*)")
			arg[k]=v
		end
		local worksheet = wb:add_worksheet(arg.name)
		if sheet.layout then
			for _,line in ipairs(sheet.layout) do
				local tag, k, v = line:match "(%w)(%d+):(.*)"
				if tag == 'c' then
					k = tonumber(k)
					worksheet:set_column(k,k, tonumber(v))
				elseif tag == 'r' then
					worksheet:set_row(tonumber(k), tonumber(v))
				end
			end
		end

		for _,v in ipairs(sheet.data) do
			local k,v = v:match("^(%w+):(.*)")
			local row, col = util.cell_to_rowcol(k)
			local number = tonumber(v)
			local frow = fill[row+1]
			local fm = styles[frow and frow[col+1]]
			if number then
				worksheet:write_number(row, col , number, fm)
			elseif v:byte() == 61 then	-- =
				worksheet:write_formula(row, col, unescape_value(v), fm, fvalue[k])
			else
				worksheet:write_string(row, col, unescape_value(v), fm)
			end
		end
	end
	wb:close()
end

local mode, f1, f2 = ...
local MODE = {}

function MODE.x2h(f1,f2)
	if f2 == nil then
		f2 = f1:match"(.*)%.xlsx$" .. ".hfs"
	end
	x2h(f1,f2)
end

function MODE.h2x(f1,f2)
	if f2 == nil then
		f2 = f1:match"(.*)%.hfs$" .. ".xlsx"
	end
	h2x(f1,f2)
end

--------------------------------------------------

local TMP_PATH = os.getenv "hfs_tmp" or ""

local function log(...)
	print(os.date(), ...)
end

local function monitor(name)
	local winapi = require "winapi"
	local ti = winapi.LastWriteTime(name)
	coroutine.yield(ti)
	while true do
		winapi.Sleep(3000)	-- sleep 3 s
		local f = io.open(name, "r+")
		if f then
			f:close()
			return
		end

		local newt = winapi.LastWriteTime(name)
		if newt ~= ti then
			ti = newt
			coroutine.yield(true)
		end
	end
end

function MODE.monitor(filename)
	local winapi = require "winapi"
	assert(filename, "Need an hfs filename")
	local fullpath, name_noext, ext = filename:match("(.*)\\(.*%.)(%w+)$")
	assert(ext:lower() == "hfs", "Only support .hfs file")
	local tmpdir = base64.hex(base64.hashkey(filename))
	tmpdir = TMP_PATH .. tmpdir
	log("Create Dir", tmpdir)
	winapi.CreateDirectory(tmpdir)
	local tmpname = tmpdir .. "\\" .. name_noext .. "xlsx"
	log(string.format("Convert %s to %s", filename, tmpname))
	h2x(filename, tmpname)
	assert(winapi.ShellExecute(tmpname), "open " .. tmpname .. " Failed")
	local mo = coroutine.wrap(monitor)
	local ft = mo(tmpname)
	while mo() do
		log(string.format("Convert back %s to %s", tmpname, filename))
		x2h(tmpname, filename)
	end
	if winapi.LastWriteTime(tmpname) ~= ft then
		log(string.format("Convert back %s to %s", tmpname, filename))
		x2h(tmpname, filename)
	end
	log("Remove ", tmpname)
	assert(winapi.DeleteFile(tmpname), "Remove " .. tmpname .. "failed")
	log("Remove Dir", tmpdir)
	assert(winapi.RemoveDirectory(tmpdir), "Remove tmp dir failed")
end

local function alert(err)
	local ok, winapi = pcall(require, "winapi")
	if ok then
		winapi.MessageBox(err)
	else
		print(err)
	end
end

local function main(mode, f1, f2)
	assert(type(f1) == "string")
	local f = assert(MODE[mode], "Invalid mode")
	f(f1,f2)
end

local ok, err = pcall(main, ...)
if not ok then
	alert(err)
end

