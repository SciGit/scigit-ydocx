#!/usr/bin/env ruby
# encoding: utf-8

require 'nokogiri'
require 'htmlentities'
require 'ydocx/markup_method'
require 'roman-numerals'

module YDocx
  Style = Struct.new(:b, :u, :i, :font, :sz, :valign, :ilvl, :numid)
  class Parser
    include MarkupMethod
    attr_accessor :indecies, :images, :result, :space
    def initialize(doc, rel, rel_files)
      @doc = Nokogiri::XML.parse(doc)
      @rel = Nokogiri::XML.parse(rel)
      @rel_files = rel_files
      @style_nodes = {}
      @styles = {}
      @numbering_desc = {}
      @numbering_count = {}
      @coder = HTMLEntities.new
      @indecies = []
      @images = []
      @result = []
      @space = '&nbsp;'
      @image_path = 'images'
      @image_style = ''
      init
      if block_given?
        yield self
      end
    end
    
    def init
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
      if rpr = node.xpath('w:rPr').first
        if b = rpr.xpath('w:b').first
          style.b = b['w:val'].nil? || b['w:val'] == '1'
        end
        if i = rpr.xpath('w:i').first
          style.i = i['w:val'].nil? || i['w:val'] == '1'
        end
        if u = rpr.xpath('w:u').first
          style.u = u['w:val'] != 'none'
        end
        if valign = rpr.xpath('w:vertAlign').first
          style.valign = valign['w:val']
        end
        if font = rpr.xpath('w:rFonts').first
          style.font = font['w:ascii']
        end
        if sz = rpr.xpath('w:sz').first
          style.sz = sz['w:val'].to_i
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
          style = Style.new()
        end
        @styles[id] = apply_style(style, node)
      end
    end
    
    def parse
      if @rel_files.has_key?('styles.xml')
        style_xml = Nokogiri::XML.parse(@rel_files['styles.xml'])
        style_xml.xpath('//w:styles//w:style').each do |style|
          @style_nodes[style['w:styleId']] = style
        end
      end
      
      @style_nodes.keys.each do |id|
        compute_style(id)
      end
      
      if @rel_files.has_key?('numbering.xml')
        num_xml = Nokogiri::XML.parse(@rel_files['numbering.xml'])
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
                  :start  => lvl.xpath('w:start').first['w:val'].to_i,
                  :numFmt => lvl.xpath('w:numFmt').first['w:val'],
                  :format => lvl.xpath('w:lvlText').first['w:val'],
                  :isLgl  => !lvl.xpath('w:isLgl').first.nil?,
                  :style  => apply_style(Style.new(), lvl)
                }
              end
            end
          end
          num.xpath('w:lvlOverride').each do |over|
            indent_level = over['w:ilvl'].to_i
            if start_over = over.xpath('w:startOverride').first
              @numbering_desc[num_id][indent_level][:start] = start_over['w:val'].to_i
            elsif lvl = over.xpath('w:lvl').first
              @numbering_desc[num_id][indent_level] = {
                :start  => lvl.xpath('w:start').first['w:val'].to_i,
                :numFmt => lvl.xpath('w:numFmt').first['w:val'],
                :format => lvl.xpath('w:lvlText').first['w:val'],
                :isLgl  => !lvl.xpath('w:isLgl').first.nil?,
                :style  => apply_style(Style.new(), lvl)
              }
            end
          end
        end
      end
      @doc.xpath('//w:document//w:body').children.map do |node|
        case node.node_name
        when 'text'
          @result << parse_paragraph(node)
        when 'tbl'
          @result << parse_table(node)
        when 'p'
          @result << parse_paragraph(node)
        else
          # skip
        end
      end
      @result
    end
    private
    def apply_fonts(style, text)
      css = ''
      if style.font
        css += sprintf("font-family: '%s';", style.font)
      end
      if style.sz
        css += sprintf("font-size: %dpt;", style.sz / 2)
      end
      if css.empty?
        text
      else
        markup :font, text, {:style => css}
      end
    end
    def apply_align(style, text)
      if style.valign == 'subscript'
        text = markup(:sub, text)
      elsif style.valign == 'superscript'
        if text =~ /^[0-9]$/
          text = "&sup" + text + ";"
        else
          text = markup(:sup, text)
        end
      end
      text
    end
    
    def character_encode(text)
      text.force_encoding('utf-8')
      # NOTE
      # :named only for escape at Builder
      text = @coder.encode(text, :named)
      text
    end
    def escape_whitespace(text)
      prev_ws = true
      new_text = ''
      text.each_char do |c|
        if c == "\n"
          new_text += "<br />"
          prev_ws = true
        elsif c =~ /[[:space:]]/
          if prev_ws
            new_text += @space
          else
            new_text += c
          end
          prev_ws = true
        else  
          new_text += c
          prev_ws = false
        end
      end
      new_text
    end
    def parse_block(node)
      nil # default no block element
    end
    def parse_image(r)
      id = nil
      additional_namespaces = {
        'xmlns:a'   => 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'xmlns:pic' => 'http://schemas.openxmlformats.org/drawingml/2006/picture'
      }
      ns = r.namespaces.merge additional_namespaces
      [
        { # old type shape
          :attr => 'id',
          :path => 'w:pict//v:shape//v:imagedata',
          :wrap => 'w:pict//v:shape//w10:wrap',
          :type => '',
        },
        { # in anchor
          :attr => 'embed',
          :path => 'w:drawing//wp:anchor//a:graphic//a:graphicData//pic:pic//pic:blipFill//a:blip',
          :wrap => 'w:drawing//wp:anchor//wp:wrapTight',
          :type => 'wrapText',
        },
        { # stand alone
          :attr => 'embed',
          :path => 'w:drawing//a:graphic//a:graphicData//pic:pic//pic:blipFill//a:blip',
          :wrap => 'w:drawing//wp:wrapTight',
          :type => 'wrapText',
        },
      ].each do |element|
        if image = r.xpath(element[:path], ns) and !image.empty?
          if wrap = r.xpath("#{element[:wrap]}", ns).first
            # TODO
            # wrap handling (currently all wrap off)
            # wrap[element[:type]] has "bothSides", "topAndBottom" and "wrapText"
            @image_style = 'display:block;'
          end
          (id = image.first[element[:attr].to_s]) && break
        end
      end
      if id
        @rel.xpath('/').children.each do |rel|
          rel.children.each do |r|
            if r['Id'] == id and r['Target']
              target = r['Target']
              source = source_path(target)
              @images << {
                :origin => target,
                :source => source
              }
              attributes = {:src => source}
              attributes.merge!({:style => @image_style}) unless @image_style.empty?
              return markup :img, [], attributes
            end
          end
        end
      end
      nil
    end
    def source_path(target)
      source = @image_path + '/'
      if defined? Magick::Image and
         ext = File.extname(target).match(/\.wmf$/).to_a[0]
        source << File.basename(target, ext) + '.png'
      else
        source << File.basename(target)
      end
    end
    def parse_paragraph(node)
      content = []
      if block = parse_block(node)
        content << block
      else # as p
        pos = 0
        style_node = node.xpath('w:pPr//w:pStyle').first
        if style_node
          style = @styles[style_node['w:val']]
        else
          style = Style.new()
        end
        style = apply_style(style, node)
        num_id = style.numid
        indent_level = style.ilvl || 0
        unless num_id.nil?
          if @numbering_desc[num_id] && num_desc = @numbering_desc[num_id][indent_level]
            format = num_desc[:format]
            is_legal = num_desc[:isLgl]
            num_style = extend_style(style, num_desc[:style])
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
              content << parse_text(format, num_style) << @space
            end
          end
        end
        
        node.xpath('.//w:r').each do |r|
          r_style = apply_style(style, r)
          if !r.xpath('w:pict').empty? || !r.xpath('w:drawing').empty?
            content << parse_image(r)
          else
            text = ''
            r.children.each do |t|
              if t.name == 'br'
                text += "\n"
              elsif t.name == 'tab'
                text += "\t"
              elsif t.name == 't'
                text += t.text
              elsif t.name == 'sym'
                val = t['w:char'].to_i(16)
                if val >= 0xf000
                  val -= 0xf000
                end
                chr_style = r_style.dup()
                if t['w:font']
                  chr_style.font = t['w:font']
                end
                content << parse_text(text, r_style)
                content << parse_text('&#x' + val.to_s(16) + ';', chr_style, true)
                text = ''
              end
            end
            unless text.empty?
              content << parse_text(text, r_style)
            end
          end
        end
      end
      content.compact!
      markup :p, content
    end
    def parse_table(node)
      table = markup :table
      
      vmerge_type = {}
      # first, compute rowspans
      node.xpath('w:tr').each_with_index do |tr, row|
        vmerge_type[row] = {}
        col = 0
        tr.xpath('w:tc').each do |tc|
          tc.xpath('w:tcPr').each do |tcpr|
            cells = 1
            if span = tcpr.xpath('w:gridSpan').first
              cells = span['w:val'].to_i
            end
            if merge = tcpr.xpath('w:vMerge').first
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
        cells = markup :tr
        col = 0
        tr.xpath('w:tc').each do |tc|
          attributes = {}
          show = true
          columns = 1
          tc.xpath('w:tcPr').each do |tcpr|
            if span = tcpr.xpath('w:gridSpan').first
              columns = attributes[:colspan] = span['w:val'].to_i
            end
            if vmerge_type[row][col] == 2
              nrow = row + 1
              while !vmerge_type[nrow].nil? and vmerge_type[nrow][col] == 1
                nrow += 1
              end
             attributes[:rowspan] = nrow - row
            end
            if align = tcpr.xpath('w:vAlign').first
              attributes[:valign] = align['w:val']
            else
              attributes[:valign] = 'top'
            end
          end
          cell = markup :td, [], attributes
          tc.xpath('w:p').each do |p|
            cell[:content] << parse_paragraph(p)
          end
          if vmerge_type[row][col] != 1
            cells[:content] << cell
          end
          col += columns
        end
        table[:content] << cells
      end
      table
    end
    def parse_text(text, style, raw = false)
      unless raw
        text = character_encode(text)
        text = escape_whitespace(text)
      end
      text = apply_fonts(style, text)
      text = apply_align(style, text)
      if style.u
        text = markup(:span, text, {:style => "text-decoration:underline;"})
      end
      if style.i
        text = markup(:em, text)
      end
      if style.b
        text = markup(:strong, text)
      end
      text
    end
  end
end
