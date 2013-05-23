#!/usr/bin/env ruby
# encoding: utf-8

require 'nokogiri'
require 'htmlentities'
require 'ydocx/markup_method'
require 'roman-numerals'
require 'rmagick'

module YDocx
  Style = Struct.new(:b, :u, :i, :strike, :caps, :smallCaps, :font, :sz, :color, :valign, :ilvl, :numid)

  class DocumentElement
   private
    attr_accessor :hash
   public
    include MarkupMethod
    def ==(elem)
      hash == elem.hash
    end
    def eql?(elem)
      hash == elem.hash
    end
    def reset_hash
      @hash = nil
    end
    def hash
      if @hash.nil?
        vals = []
        instance_variables.each do |var|
          if var != "@hash"
            vals << instance_variable_get(var)
          end
        end
        @hash = vals.hash
      end
      @hash
    end
  end
  
  class Image < DocumentElement
    attr_accessor :height, :width, :src, :wrap, :img_hash
    def initialize(height=nil, width=nil, src=nil, wrap='inline')
      @height = height
      @width = width
      @src = src
      @wrap = wrap
    end
    def to_markup
      attributes = {}
      style = ['display: ' + @wrap]
      if @height
        style << "height: #{@height}"
        style << "width: #{@width}"
      end
      attributes[:style] = style.join('; ')
      if @src
        attributes[:src] = @src
      else
        attributes[:alt] = 'Unknown image'
      end
      markup :img, '', attributes
    end
  end
  
  class Run < DocumentElement
    attr_accessor :text, :style
    def initialize(text='', style=Style.new)
      @text = text
      @style = style
    end
    def to_markup
      css = []
      text = @text
      if @style.font
        css << sprintf("font-family: '%s'", @style.font)
      end
      if @style.sz
        css << sprintf("font-size: %dpt", @style.sz / 2)
      end
      if @style.color
        css << sprintf("color: #%s", @style.color)
      end
      if @style.u
        css << 'text-decoration: underline'
      end
      if @style.i
        css << 'font-style: italic'
      end
      if @style.b
        css << 'font-weight: bold'
      end
      if @style.caps
        css << 'text-transform: uppercase'
      end
      if @style.smallCaps
        css << 'font-variant: small-caps'
      end
      if @style.strike
        if @style.u
          text = markup :span, text, {:style => "text-decoration: line-through"}
        else
          css << 'text-decoration: line-through'
        end
      end
      if @style.valign == 'subscript'
        text = markup :sub, text
      elsif @style.valign == 'superscript'
        text = markup :sup, text
      end
      if css.empty?
        text
      else
        markup :span, text, {:style => css.join("; ")}
      end
    end
    def hash
      [@text, @style.hash].hash
    end
  end
  
  class Paragraph < DocumentElement
    attr_accessor :runs, :align
    def initialize(align='left')
      @align = align
      @runs = []
    end
    def merge_runs
      # merge adjacent runs with identical formatting.
      @new_runs = []
      cur_text = ''
      cur_style = Style.new
      i = 0
      while i < @runs.length
        if !@runs[i].is_a?(Run)
          @new_runs << @runs[i]
          i += 1
        else
          j = i + 1
          text = @runs[i].text.dup
          while j < @runs.length && @runs[j].is_a?(Run) && (@runs[j].style == @runs[i].style || @runs[j].text == '<br />')
            text << @runs[j].text
            j += 1
          end
          @new_runs << Run.new(text, @runs[i].style)
          i = j
        end
      end
      @runs = @new_runs
    end
    def to_markup
      res = []
      @runs.each do |run|
        res << run.to_markup
      end
      markup :p, res, {:align => @align}
    end
    def hash
      [@runs.hash, @align].hash
    end
  end
  
  class Cell < DocumentElement
    attr_accessor :rowspan, :colspan, :height, :width, :valign, :blocks
    def initialize(rowspan=1, colspan=1, height=nil, width=nil, valign='top')
      @rowspan = rowspan
      @colspan = colspan
      @height = height
      @width = width
      @valign = valign
      @blocks = []
    end
    def to_markup
      contents = []
      @blocks.each do |block|
        contents << block.to_markup
      end
      css = ["vertical-align: #{@valign}"]
      if height
        css << "height: #{@height}"
      end
      if width
        css << "width: #{@width}"
      end
      markup :td, contents, {
        :rowspan => @rowspan,
        :colspan => @colspan,
        :style => css.join('; ')
      }
    end
  end
  
  class Table < DocumentElement
    attr_accessor :cells
    def initialize
      @cells = []
    end
    def to_markup
      rows = []
      @cells.each do |row|
        cells = []
        row.each do |cell|
          cells << cell.to_markup
        end
        rows << markup(:tr, cells)
      end
      markup :table, rows
    end
  end
  
  class ParsedDocument < DocumentElement
    attr_accessor :blocks
    def initialize
      @blocks = []
    end
    def to_markup
      body = []
      @blocks.each do |block|
        body << block.to_markup
      end
      body
    end
  end
  
  class Parser
    attr_accessor :images, :result, :space
    def initialize(doc, rel, rel_files)
      @doc = Nokogiri::XML.parse(doc)
      @rel = Nokogiri::XML.parse(rel)
      @rel_files = rel_files
      @style_nodes = {}
      @styles = {}
      @theme_fonts = {}
      @numbering_desc = {}
      @numbering_count = {}
      @coder = HTMLEntities.new
      @images = []
      @result = ParsedDocument.new
      @image_path = 'images'
      @image_style = ''
      init
      if block_given?
        yield self
      end
    end
    
    def init
    end
    
    def get_bool(node)
      return !node.nil? && (node['w:val'].nil? || node['w:val'] == 'true' || node['w:val'] == '1')
    end
    
    def extend_style(old_style, new_style)
      style = Style.new()
      new_style.members.each do |key|
        style[key] = new_style[key] || old_style[key]
      end
      style
    end
    
    def apply_style(old_style, node)
      style = old_style.dup()
      
      if numpr = node.at_xpath('w:pPr//w:numPr')
        if ilvl = numpr.at_xpath('w:ilvl')
          style.ilvl = ilvl['w:val'].to_i
        end
        if numid = numpr.at_xpath('w:numId')
          style.numid = numid['w:val'].to_i
        end
      end
      
      if rpr = node.at_xpath('w:rPr')
        if rstyle = rpr.at_xpath('w:rStyle')
          style = extend_style(style, @styles[rstyle['w:val']])
        end
        [:b, :i, :strike, :caps, :smallCaps].each do |mod|
          if modnode = rpr.at_xpath('w:' + mod.to_s)
            style[mod] = get_bool(modnode)
          end
        end
        if !style.strike && (dstrike = rpr.at_xpath('w:dstrike'))
          style.strike = get_bool(dstrike)
        end
        if u = rpr.at_xpath('w:u')
          style.u = u['w:val'] != 'none' # TODO there are other types of underlines
        end
        if valign = rpr.at_xpath('w:vertAlign')
          style.valign = valign['w:val']
        end
        if font = rpr.at_xpath('w:rFonts')
          if !font['w:ascii'].nil?
            style.font = font['w:ascii']
          elsif !font['w:asciiTheme'].nil?
            theme = font['w:asciiTheme'][0, 5]
            style.font = @theme_fonts[theme]
          end
        end
        if sz = rpr.at_xpath('w:sz')
          style.sz = sz['w:val'].to_i
        end
        if color = rpr.at_xpath('w:color')
          style.color = color['w:val']
        end
      end
      style
    end
    
    def compute_style(id)
      node = @style_nodes[id]
      if @styles.has_key?(id)
        @styles[id]
      else
        if based = node.at_xpath('w:basedOn')
          style = compute_style(based['w:val'])
        else
          style = @default_style
        end
        @styles[id] = apply_style(style, node)
      end
    end
    
    def parse
      if theme_file = @rel_files.select { |file| file[:type] =~ /relationships\/theme$/ }.first
        theme_xml = Nokogiri::XML.parse(theme_file[:stream])
        ['major', 'minor'].each do |type|
          if font = theme_xml.at_xpath(".//a:#{type}Font//a:latin")
            @theme_fonts[type] = font['typeface']
          end
        end
      end
      
      if style_file = @rel_files.select { |file| file[:type] =~ /relationships\/styles$/ }.first
        style_xml = Nokogiri::XML.parse(style_file[:stream])
        style_xml.xpath('//w:styles//w:style').each do |style|
          @style_nodes[style['w:styleId']] = style
        end
        @default_style = Style.new()
        if def_style = style_xml.at_xpath('//w:styles//w:docDefaults//w:rPrDefault')
          @default_style = apply_style(Style.new(), def_style)
        end
      end
      
      @style_nodes.keys.each do |id|
        compute_style(id)
      end
      
      if num_file = @rel_files.select { |file| file[:type] =~ /relationships\/numbering$/ }.first
        num_xml = Nokogiri::XML.parse(num_file[:stream])
        abstract_nums = {}
        num_xml.xpath('//w:numbering//w:abstractNum').each do |abstr|
          abstract_nums[abstr['w:abstractNumId']] = abstr
        end
        num_xml.xpath('//w:numbering//w:num').each do |num|
          num_id = num['w:numId'].to_i
          @numbering_desc[num_id] = {}
          @numbering_count[num_id] = {}
          num.xpath('w:abstractNumId').each do |abstr|
            if abstract_nums.has_key?(abstr['w:val'])
              abstract_nums[abstr['w:val']].xpath('w:lvl').each do |lvl|
                indent_level = lvl['w:ilvl'].to_i
                @numbering_count[num_id][indent_level] = 0
                @numbering_desc[num_id][indent_level] = {
                  :start  => lvl.at_xpath('w:start')['w:val'].to_i,
                  :numFmt => lvl.at_xpath('w:numFmt')['w:val'],
                  :format => lvl.at_xpath('w:lvlText')['w:val'],
                  :isLgl  => get_bool(lvl.at_xpath('w:isLgl')),
                  :style  => apply_style(Style.new(), lvl)
                }
              end
            end
          end
          num.xpath('w:lvlOverride').each do |over|
            indent_level = over['w:ilvl'].to_i
            if start_over = over.at_xpath('w:startOverride')
              @numbering_desc[num_id][indent_level][:start] = start_over['w:val'].to_i
            elsif lvl = over.at_xpath('w:lvl')
              @numbering_desc[num_id][indent_level] = {
                :start  => lvl.at_xpath('w:start')['w:val'].to_i,
                :numFmt => lvl.at_xpath('w:numFmt')['w:val'],
                :format => lvl.at_xpath('w:lvlText')['w:val'],
                :isLgl  => get_bool(lvl.at_xpath('w:isLgl')),
                :style  => apply_style(Style.new(), lvl)
              }
            end
          end
        end
      end
      @doc.xpath('//w:document//w:body').children.map do |node|
        case node.node_name
        when 'text'
          @result.blocks << parse_paragraph(node)
        when 'tbl'
          @result.blocks << parse_table(node)
        when 'p'
          @result.blocks << parse_paragraph(node)
        else
          # skip
        end
      end
      @result
    end
    
   private
    def character_encode(text)
      text.force_encoding('utf-8')
      # NOTE
      # :named only for escape at Builder
      text = @coder.encode(text, :named)
      text
    end
    def parse_image(r)
      id = nil
      img = Image.new
      additional_namespaces = {
        'xmlns:a'   => 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'xmlns:pic' => 'http://schemas.openxmlformats.org/drawingml/2006/picture'
      }
      ns = r.namespaces.merge additional_namespaces
      [
        { # old type shape
          :attr => 'id',
          :path => './/w:pict//v:shape//v:imagedata',
        },
        { # in anchor
          :attr => 'r:embed',
          :path => './/w:drawing//wp:anchor',
        },
        { # inline
          :attr => 'r:embed',
          :path => './/w:drawing//wp:inline',
        },
      ].each do |element|
        if image = r.at_xpath(element[:path], ns)
          if wrap = image.at_xpath('wp:wrapTopAndBottom', ns)
            img.wrap = 'block'
          end
          if size = image.at_xpath('wp:extent', ns)
            img.width = size['cx'].to_i / 9525
            img.height = size['cy'].to_i / 9525
          end
          if blip = image.at_xpath('a:graphic//a:graphicData//pic:pic//pic:blipFill//a:blip', ns)
            image = blip
          end
          id = image[element[:attr]]              
          if id
            if file = @rel_files.select{ |file| file[:id] == id }.first
              target = file[:target]
              source = source_path(target)
              data = file[:stream].read
              @images << {
                :origin => target,
                :source => source,
                :data => data,
              }
              img.src = source
              img.img_hash = data.hash
            end
          else
            img.img_hash = image.to_s.hash
          end
          break
        end
      end
      img
    end
    def source_path(target)
      source = @image_path + '/'
      if defined? Magick::Image and
         ext = File.extname(target).match(/\.(w|e)mf$/).to_a[0] # EMF may not work outside of windows!
        source << File.basename(target, ext) + '.png'
      else
        source << File.basename(target)
      end
    end
    def parse_paragraph(node)
      paragraph = Paragraph.new
      if style_node = node.at_xpath('w:pPr//w:pStyle')
        style = @styles[style_node['w:val']]
      else
        style = @default_style
      end
      style = apply_style(style, node)
      num_id = style.numid
      indent_level = style.ilvl || 0
      unless num_id.nil?
        if @numbering_desc[num_id] && num_desc = @numbering_desc[num_id][indent_level]
          format = num_desc[:format]
          is_legal = num_desc[:isLgl]
          num_style = style.dup()
          # It seems that text size from pPr.rPr applies to numbering in some cases...
          if sz = node.at_xpath('w:pPr//w:rPr//w:sz')
            num_style[:sz] = sz['w:val'].to_i
          end
          num_style = extend_style(num_style, num_desc[:style])
          for ilvl in 0..indent_level
            if num_desc = @numbering_desc[num_id][ilvl]
              num = num_desc[:start] + @numbering_count[num_id][ilvl] - (ilvl < indent_level ? 1 : 0)
              replace = '%' + (ilvl+1).to_s
              next if !format.include?(replace)
              str = case (is_legal and ilvl < indent_level) ? 'decimal' : num_desc[:numFmt]
              when 'decimalZero'
                sprintf("%02d", num)
              when 'upperRoman'
                RomanNumerals.to_roman(num)
              when 'lowerRoman'
                RomanNumerals.to_roman(num).downcase
              when 'upperLetter'
                letter = (num-1) % 26
                rep = (num-1) / 26 + 1
                (letter + 65).chr * rep
              when 'lowerLetter'
                letter = (num-1) % 26
                rep = (num-1) / 26 + 1
                (letter + 97).chr * rep
              # todo: idk
              # when 'ordinal'
              # when 'cardinalText'
              # when 'ordinalText'
              when 'bullet'
                '&bull;'
              else
                num.to_s
              end
            end
            format = format.sub(replace, str)
          end              
          @numbering_count[num_id][indent_level] += 1
          # reset higher counts
          @numbering_count[num_id].each_key do |level|
            if level > indent_level
              @numbering_count[num_id][level] = 0
            end
          end
          unless format == ''
            paragraph.runs << parse_text(format + ' ', num_style)
          end
        end
      end
      
      node.children.each do |child|
        if !child.xpath('.//w:pict').empty? || !child.xpath('.//w:drawing').empty?
          paragraph.runs << parse_image(child)
          next
        end
        # take care of things like smarttags (which contain runs)
        runs = child.xpath('.//w:r')
        if child.name == 'r'
          runs << child
        end
        runs.each do |r|
          r_style = apply_style(style, r)
          text = ''
          r.children.each do |t|
            if t.name == 'br'
              text += "\n"
            elsif t.name == 'tab'
              text += "        "
            elsif t.name == 't'
              text += t.text
            elsif t.name == 'sym'
              text += t.text
              val = t['w:char'].to_i(16)
              if val >= 0xf000
                val -= 0xf000
              end
              chr_style = r_style.dup()
              if t['w:font']
                chr_style.font = t['w:font']
              end
              paragraph.runs << parse_text(text, r_style)
              paragraph.runs << parse_text('&#x' + val.to_s(16) + ';', chr_style, true)
              text = ''
            end
          end
          unless text.empty?
            paragraph.runs << parse_text(text, r_style)
          end
        end
      end
      if jc = node.at_xpath('w:pPr//w:jc')
        paragraph.align = jc['w:val']
        if paragraph.align == 'both'
          paragraph.align = 'justify'
        end
      end
      paragraph.merge_runs
      paragraph
    end
    def parse_table(node)
      table = Table.new
      
      vmerge_type = {}
      # first, compute rowspans
      node.xpath('w:tr').each_with_index do |tr, row|
        vmerge_type[row] = {}
        col = 0
        tr.xpath('w:tc').each do |tc|
          tc.xpath('w:tcPr').each do |tcpr|
            cells = 1
            if span = tcpr.at_xpath('w:gridSpan')
              cells = span['w:val'].to_i
            end
            if merge = tcpr.at_xpath('w:vMerge')
              if merge['w:val'].nil?
                vmerge_type[row][col] = 1;
              else
                vmerge_type[row][col] = 2
              end
            else
              vmerge_type[row][col] = 0
            end
            col += cells
          end
        end
      end
      
      node.xpath('w:tr').each_with_index do |tr, row|
        row_height = nil
        if trh = tr.at_xpath('w:trPr//w:trHeight')
          row_height = trh['w:val'].to_i * 96 / 1440
        end
        table.cells << []
        col = 0
        tr.xpath('w:tc').each do |tc|
          cell = Cell.new
          cell.height = row_height
          columns = 1
          if tcpr = tc.at_xpath('w:tcPr')
            if span = tcpr.at_xpath('w:gridSpan')
              columns = cell.colspan = span['w:val'].to_i
            end
            if w = tcpr.at_xpath('w:tcW')
              cell.width = w['w:w'].to_i * 96 / 1440
            end
            if vmerge_type[row][col] == 2
              nrow = row + 1
              while !vmerge_type[nrow].nil? and vmerge_type[nrow][col] == 1
                nrow += 1
              end
             cell.rowspan = nrow - row
            end
            if align = tcpr.at_xpath('w:vAlign')
              cell.valign = align['w:val']
            end
          end
          tc.xpath('w:p').each do |p|
            cell.blocks << parse_paragraph(p)
          end
          if vmerge_type[row][col] != 1
            table.cells[row] << cell
          end
          col += columns
        end
      end
      table
    end
    def parse_text(text, style, raw = false)
      unless raw
        text = character_encode(text)
      end
      text_style = style.dup
      text_style.ilvl = text_style.numid = nil
      Run.new text, text_style
    end
  end
end
