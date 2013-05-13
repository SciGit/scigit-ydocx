#!/usr/bin/env ruby
# encoding: utf-8

require 'nokogiri'
require 'htmlentities'
require 'ydocx/markup_method'
require 'roman-numerals'

module YDocx
  class Parser
    include MarkupMethod
    attr_accessor :indecies, :images, :result, :space
    def initialize(doc, rel, rel_files)
      @doc = Nokogiri::XML.parse(doc)
      @rel = Nokogiri::XML.parse(rel)
      @rel_files = rel_files
      @styles = {}
      @numbering_styles = {}
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
    def parse
      if @rel_files.has_key?('styles.xml')
        style_xml = Nokogiri::XML.parse(@rel_files['styles.xml'])
        style_xml.xpath('//w:styles//w:style').each do |style|
          @styles[style['w:styleId']] = style
        end
      end
      
      if @rel_files.has_key?('numbering.xml')
        num_xml = Nokogiri::XML.parse(@rel_files['numbering.xml'])
        abstract_nums = {}
        num_xml.xpath('//w:numbering//w:abstractNum').each do |abstr|
          abstract_nums[abstr['w:abstractNumId']] = abstr
        end
        num_xml.xpath('//w:numbering//w:num').each do |num|
          num_id = num['w:numId'].to_i
          @numbering_styles[num_id] = {}
          @numbering_count[num_id] = {}
          num.xpath('w:abstractNumId').each do |abstr|
            if abstract_nums.has_key?(abstr['w:val'])
              abstract_nums[abstr['w:val']].xpath('w:lvl').each do |lvl|
                indent_level = lvl['w:ilvl'].to_i
                @numbering_count[num_id][indent_level] = 0
                @numbering_styles[num_id][indent_level] = {
                  :start  => lvl.xpath('w:start').first['w:val'].to_i,
                  :numFmt => lvl.xpath('w:numFmt').first['w:val'],
                  :format => lvl.xpath('w:lvlText').first['w:val'],
                  :isLgl  => !lvl.xpath('w:isLgl').first.nil?
                }
              end
            end
          end
          num.xpath('w:lvlOverride').each do |over|
            indent_level = over['w:ilvl'].to_i
            if start_over = over.xpath('w:startOverride').first
              @numbering_styles[num_id][indent_level][:start] = start_over['w:val'].to_i
            elsif lvl = over.xpath('w:lvl').first
              @numbering_styles[num_id][indent_level] = {
                :start  => lvl.xpath('w:start').first['w:val'].to_i,
                :numFmt => lvl.xpath('w:numFmt').first['w:val'],
                :format => lvl.xpath('w:lvlText').first['w:val'],
                :isLgl  => !lvl.xpath('w:isLgl').first.nil?
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
    def apply_fonts(rpr, text)
      symbol = false
      unless rpr.xpath('w:rFonts').empty?
        rpr.xpath('w:rFonts').each do |font|
          if font.values.include? 'Symbol'
            symbol = true
          end
          break if symbol
        end
      end
      if symbol
        _text = ''
        text.unpack('U*').each do |char|
          _text << character_replace(char.to_s(16))
        end
        text = _text
      end
      text
    end
    def apply_align(rpr, text)
      unless rpr.xpath('w:vertAlign').empty?
        script = rpr.xpath('w:vertAlign').first['val'].to_sym
        if script == :subscript
          text = markup(:sub, text)
        elsif script == :superscript
          if text =~ /^[0-9]$/
            text = "&sup" + text + ";"
          else
            text = markup(:sup, text)
          end
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
    def character_replace(code)
      code = '0x' + code
      # NOTE
      # replace with rsemble html character ref
      # Symbol Font to HTML Character named ref
      case code
      when '0xf020' # '61472'
        ""
      when '0xf025' # '61477'
        "%"
      when '0xf02b' # '61482'
        "*"
      when '0xf02b' # '61483'
        "+"
      when '0xf02d' # '61485'
        "-"
      when '0xf02f' # '61487'
        "/"
      when '0xf03c' # '61500'
        "&lt;"
      when '0xf03d' # '61501'
        "="
      when '0xf03e' # '61502'
        "&gt;"
      when '0xf040' # '61504'
        "&cong;"
      when '0xf068' # '61544'
        "&eta;"
      when '0xf071' # '61553'
        "&theta;"
      when '0xf06d' # '61549'
        "&mu;"
      when '0xf0a3' # '61603'
        "&le;"
      when '0xf0ab' # '61611'
        "&harr;"
      when '0xf0ac' # '61612'
        "&larr;"
      when '0xf0ad' # '61613'
        "&uarr;"
      when '0xf0ae' # '61614'
        "&rarr;"
      when '0xf0ad' # '61615'
        "&darr;"
      when '0xf0b1' # '61617'
        "&plusmn;"
      when '0xf0b2' # '61618'
        "&Prime;"
      when '0xf0b3' # '61619'
        "&ge;"
      when '0xf0b4' # '61620'
        "&times;"
      when '0xf0b7' # '61623'
        "&sdot;"
      else
        #p "code : " + ("&#%s;" % code)
        #p "hex  : " + code.hex.to_s
        #p "char : " + @coder.decode("&#%s;" % code.hex.to_s)
      end
    end
    def optional_escape(text)
      text
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
    def get_numbering(node)
      if num_prop = node.xpath('w:pPr//w:numPr').first      
        if (base = node.xpath('w:basedOn').first) && (base_style = @styles[base['w:val']])
          indent_level, num_id = get_numbering(base_style)          
        end
        if ilvl = num_prop.xpath('w:ilvl').first
          indent_level = ilvl['w:val'].to_i
        end
        if nid = num_prop.xpath('w:numId').first
          num_id = nid['w:val'].to_i
        end
      end
      return indent_level, num_id
    end
    def parse_paragraph(node)
      content = []
      if block = parse_block(node)
        content << block
      else # as p
        pos = 0
        style = node.xpath('w:pPr//w:pStyle').first
        if style
          style = @styles[style['w:val']]
        end
        indent_level, num_id = get_numbering(node)
        if num_id.nil? && style
          indent_level, num_id = get_numbering(style)
        end
        indent_level ||= 0
        unless num_id.nil?
          if @numbering_styles[num_id] && style = @numbering_styles[num_id][indent_level]
            format = style[:format]
            is_legal = style[:isLgl]
            for ilvl in 0..indent_level
              if style = @numbering_styles[num_id][ilvl]
                num = style[:start] + @numbering_count[num_id][ilvl] - (ilvl < indent_level ? 1 : 0)
                replace = '%' + (ilvl+1).to_s
                next if !format.include?(replace)
                str = case (is_legal and ilvl < indent_level) ? 'decimal' : style[:numFmt]
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
            content << format + ' ' unless format == ''
          end
        end
        node.xpath('w:r').each do |r|
          unless r.xpath('w:br').empty?
            content << "<br />"
          end
          unless r.xpath('w:t').empty?
            content << parse_text(r, (pos == 0)) # rm indent
            pos += 1
          else
            unless r.xpath('w:tab').empty?
              if content.last != @space and pos != 0 # ignore tab at line head
                content << @space
                pos += 1
              end
            end
            unless r.xpath('w:sym').empty?
              code = r.xpath('w:sym').first['w:char'].downcase # w:char
              content << character_replace(code)
              pos += 1
            end
            if !r.xpath('w:pict').empty? or !r.xpath('w:drawing').empty?
              content << parse_image(r)
            end
          end
        end
      end
      content.compact!
      unless content.empty?
        paragraph = content.select do |c|
          c.is_a?(Hash) and c[:tag].to_s =~ /^h[1-9]/u
        end.empty?
        if paragraph
          markup :p, content
        else
          content.first
        end
      else
        {}
      end
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
    def parse_text(r, lstrip=false)
      text = r.xpath('w:t').map(&:text).join('')
      text = character_encode(text)
      text = optional_escape(text)
      text = text.lstrip if lstrip
      if rpr = r.xpath('w:rPr')
        text = apply_fonts(rpr, text)
        text = apply_align(rpr, text)
        unless rpr.xpath('w:u').empty? || rpr.xpath('w:u').first['w:val'] == '0'
          text = markup(:span, text, {:style => "text-decoration:underline;"})
        end
        unless rpr.xpath('w:i').empty? || rpr.xpath('w:i').first['w:val'] == '0'
          text = markup(:em, text)
        end
        unless rpr.xpath('w:b').empty? || rpr.xpath('w:b').first['w:val'] == '0'
          text = markup(:strong, text)
        end
      end
      text
    end
  end
end
