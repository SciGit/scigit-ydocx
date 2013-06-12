require 'ydocx/markup_method'

module YDocx
  class Hashable
   private
    @@ignored_variables = {:@src => 1, :@css_class => 1, :@parent => 1}
   public
    attr_accessor :cached_hash
    def ==(elem)
      hash == elem.hash
    end
    def !=(elem)
      hash != elem.hash
    end
    def eql?(elem)
      hash == elem.hash
    end
    def reset_hash
      @hash = nil
    end
    def self.murmur_hash(obj)
      if obj.is_a?(Hashable) && !obj.cached_hash.nil?
        return obj.cached_hash
      end
      hash = MurmurHash3::Native32.murmur3_32_str_hash(obj.class.name)
      if obj.is_a?(Hashable)
        obj.instance_variables.each do |var|
          if var != :@cached_hash && !@@ignored_variables.include?(var)
            hash = MurmurHash3::Native32.murmur3_32_int32_hash(
                murmur_hash(obj.instance_variable_get(var)), hash)
          end
        end
        obj.cached_hash = hash
      elsif obj.is_a?(Array)
        obj.each do |x|
          hash = MurmurHash3::Native32.murmur3_32_int32_hash(murmur_hash(x), hash)
        end
      elsif obj.nil?
        # just hashing the class name will do
      else
        hash = MurmurHash3::Native32.murmur3_32_str_hash(obj.to_s, hash)
      end
      hash
    end
    def hash
      Hashable.murmur_hash(self)
    end
  end

  class Style < Hashable
    attr_accessor :b, :u, :i, :strike, :caps, :smallCaps, :font, :sz, :color, :valign, :ilvl, :numid
    def apply(new_style)
      new_style.instance_variables.each do |key|
        instance_variable_set(key, new_style.instance_variable_get(key))
      end
      self
    end
  end

  class DocumentElement < Hashable
    include MarkupMethod
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
        style << "height: #{@height}px"
        style << "width: #{@width}px"
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
    def length
      text.length
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
    def to_s
      @text
    end
  end
  
  class RunGroup < DocumentElement
    attr_accessor :runs, :css_class
    def initialize(runs = [])
      @runs = runs
    end
    def length
      @runs.map(&:length).reduce(0, :+)
    end      
    def self.get_type(char)
      if char == ' '
        :space
      elsif char == "\n"
        :newline
      else
        :word
      end
    end
    def self.merge_runs(runs)
      # merge adjacent runs with identical formatting.
      new_runs = []
      cur_text = ''
      cur_style = Style.new
      i = 0
      while i < runs.length
        if !runs[i].is_a?(Run)
          new_runs << runs[i]
          i += 1
        else
          j = i + 1
          text = runs[i].text.dup
          while j < runs.length && runs[j].is_a?(Run) && (runs[j].style == runs[i].style || runs[j].text == "\n")
            text << runs[j].text
            j += 1
          end
          new_runs << Run.new(text, runs[i].style)
          i = j
        end
      end
      new_runs
    end
    def self.split_runs(runs)
      # get chunks of words, newlines, and contiguous spaces; each chunk may have multiple runs
      cur_group = RunGroup.new
      cur_text = ''
      cur_type = nil
      groups = []
      runs.each do |run|
        if run.is_a? Run
          run.text.each_char do |c|
            if cur_type.nil? || (cur_type != :newline && get_type(c) == cur_type)
              cur_text += c
            else
              cur_group.runs << Run.new(cur_text, (cur_type == :newline ? Style.new : run.style)) unless cur_text.empty?
              unless cur_group.runs.empty?
                groups << cur_group
                cur_group = RunGroup.new
              end
              cur_text = c
            end
            
            cur_type = get_type(c)
          end
        elsif run.is_a? Image
          unless cur_group.runs.empty?
            groups << cur_group
            cur_group = RunGroup.new
          end
          groups << RunGroup.new([run])
        end
        unless cur_text.empty?
          cur_group.runs << Run.new(cur_text, run.style)
          cur_text = ''
        end
      end
      groups << cur_group unless cur_group.runs.empty?
      groups
    end
    def to_markup
      classes = []
      classes << @css_class if @css_class
      # Image spans need to be blocks to size correctly
      classes << 'block' unless @runs.find_index{ |r| r.is_a?(Image) }.nil?
      markup :span, @runs.map { |r| r.to_markup }, {:class => classes.join(' ')}
    end
  end
  
  class Paragraph < DocumentElement
    attr_accessor :groups, :align
    def initialize(align='left')
      @align = align
      @groups = []
    end
    def length
      @groups.map(&:length).reduce(0, :+)
    end
    def get_chunks
      @groups.map { |g| g.runs }
    end
    def to_markup
      runs = @groups.map { |g| g.css_class ? g : g.runs }.flatten
      markup :p, RunGroup.merge_runs(runs).map(&:to_markup), {:align => @align}
    end
  end
  
  class Cell < DocumentElement
    attr_accessor :rowspan, :colspan, :height, :width, :valign, :blocks
    attr_accessor :css_class, :row, :col, :parent
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
        css << "height: #{@height}px"
      end
      if width
        css << "width: #{@width}px"
      end
      markup :td, contents, {
        :class => @css_class,
        :rowspan => @rowspan,
        :colspan => @colspan,
        :style => css.join('; ')
      }
    end
  end
  
  class Row < DocumentElement
    attr_accessor :cells
    def initialize  
      @cells = []
    end
    def to_markup
      markup :tr, @cells.map { |c| c.to_markup }
    end
    def get_chunks
      @cells.map { |c| [c] }
    end
  end
  
  class Table < DocumentElement
    attr_accessor :rows
    def initialize
      @rows = []
    end
    def to_markup
      markup :p, (markup :table, @rows.map { |r| r.to_markup })
    end
    def get_chunks
      @rows.map { |c| c.get_chunks }.reduce(:+)
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
end
