#!/usr/bin/env ruby
# encoding: utf-8

require 'nokogiri'
require 'erb'
require 'ostruct'
require 'andand'
require 'money'

class CheckboxValue
  attr_accessor :value

  CHECKED_CHAR = "\u00fe"
  UNCHECKED_CHAR = "\u00a8"

  def initialize(name, value)
    @value = to_bool(value)
    @str = (@value ? CHECKED_CHAR : UNCHECKED_CHAR)  + ' ' + name
  end

  def to_s
    @str
  end

  def !
    !@value
  end
end

class Array
  def second
    self[1]
  end
end

def to_bool(value)
  if value.nil?
    return false
  end
  if value.is_a?(CheckboxValue)
    return value.value
  end
  value = value.to_s.downcase
  return !value.empty? && value != 'no' && value != 'false' && value != 'null'
end

module YDocx
  DELETE_TEXT = '###DELETE_ME###'
  SECTION_TEXT = /<%\s*#\s*section_(start|end)\s*([^>]*)%>/

  class ErbBinding < OpenStruct
    attr_accessor :_replace

    def render(template)
      ERB.new(template).result(binding)
    rescue Exception => e
      raise "Could not render: " + template + "\n" + e.to_s
    end

    def eval(x)
      binding.eval(x)
    rescue Exception => e
      raise "Could not evaluate: " + x + "\n" + e.to_s
    end

    def showif(x)
      return to_bool(x) ? '' : DELETE_TEXT
    end

    def selfif(x)
      return (to_bool(x) && x) || DELETE_TEXT
    end

    def hide
      DELETE_TEXT
    end

    # This will be handled in other code.
    def ifblock(x)
      ''
    end

    def endblock(x = nil)
      ''
    end

    def append(x, y, placeholder = '')
      return x && x.to_s + y || placeholder
    end

    def date(x)
      if x.to_i > 0
        return Time.at(x.to_i).strftime('%Y/%-m/%-d')
      else
        return nil
      end
    end

    def currency(x)
      if x.nil?
        nil
      elsif x == x.to_i.to_s
        Money.new(x.to_i * 100).format :no_cents
      elsif x.match /^[0-9]/
        '$' + x
      else
        x
      end
    end

    def replace(file, section = nil)
      @_replace = {:file => file, :section => section}
    end
  end

  class TemplateParser
    OPTION_DEFAULTS = {
      :placeholder => 'None or N/A'
    }

    VAR_PATTERN = /<%=(([^%]|%(?!>))*)%>/

    def initialize(doc, fields = [], options = {})
      @label_nodes = {}
      @label_root = {}
      @node_label = {}
      @doc = doc
      @fields = {}
      @options = options.merge(OPTION_DEFAULTS)
      @sections = {}
      get_fields(@fields, fields)
    end

    def get_fields(store, fields)
      fields.each do |field|
        if field[:type].is_a? Array
          substore = field[:id][-1] == 's' ? (store[field[:id]] = {}) : store
          get_fields(substore, field[:type])
        elsif field[:id]
          store[field[:id]] = field
        end
      end
    end

    def preprocess(data, fields)
      if data.is_a?(Hash) && data['value']
        return preprocess(data['value'], fields)
      end

      if !fields.andand[:id].nil?
        case fields[:type]
        when 'checkbox'
          vals = {}
          fields[:options].each do |k, v|
            vals[k] = CheckboxValue.new(v, to_bool(data.andand[k]))
          end
          return OpenStruct.new(vals)
        when 'radio'
          return fields[:options][data]
        when 'currency'
          return data && data != '' && Money.new(data.to_i * 100).format(:no_cents)
        when 'units'
          if data.nil? || data['qty'].nil? && data['qty'] != ''
            return nil
          elsif data['unit'].nil? && !data['unit'].empty?
            return data['qty']
          else
            return data['qty'] + ' ' + data['unit'].downcase
          end
        end
      elsif data.is_a?(Hash)
        data.each do |k, v|
          data[k] = preprocess(v, fields.andand[k])
        end
        fields.andand.each do |k, v|
          if data[k].nil?
            if v[:id].nil?
              data[k] = []
            else
              data[k] = preprocess(nil, v)
            end
          end
        end
        return OpenStruct.new(data)
      elsif data.is_a?(Array)
        return data.map { |d| preprocess(d, fields) }
      end

      return data
    end

    def replace(data)
      data = preprocess(data, @fields)
      @erb_binding = ErbBinding.new(data)
      @erb_binding.placeholder = @options[:placeholder]

      doc = Nokogiri::XML.parse(@doc)
      root = doc.at_xpath('//w:document//w:body')
      group_values(root)
      replace_sections(root)
      replace_runs(root, data)

      if @erb_binding._replace
        return @erb_binding._replace
      end

      remove_empty(root)
      trim_empty_paragraphs(root)

      return doc
    end

    def list_sections
      doc = Nokogiri::XML.parse(@doc)
      root = doc.at_xpath('//w:document//w:body')
      group_values(root)
      @sections.keys
    end

    # Much faster than xpath; we don't have to do any parsing etc
    def find_child(node, name)
      node.children.find { |child | child.name == name }
    end

    def group_values(node)
      last_ifblock = nil
      node.xpath('.//w:p').each do |p|
        cur_child = 0
        while cur_child < p.children.length
          r = p.children[cur_child]
          text = find_child(r, 't')
          prop = find_child(r, 'rPr')

          if text.nil?
            cur_child += 1
            next
          end

          prop = prop ? prop.to_s : ''
          while cur_child + 1 < p.children.length
            next_r = p.children[cur_child + 1]
            next_text = find_child(next_r, 't')
            next_prop = find_child(next_r, 'rPr')
            next_prop = next_prop ? next_prop.to_s : ''

            if next_r.name.start_with?('bookmark') || next_r.name == 'proofErr'
              next_r.remove # These are inserted randomly, can't really tell what they do
            elsif next_prop == prop && next_text
              whitespace = false
              next_r.children.each do |c|
                if c.name == 'br' || c.name == 'tab'
                  whitespace = true
                  break
                elsif c == next_text
                  break
                end
              end
              if whitespace
                break
              end

              text.content += next_text.content
              if next_text['xml:space']
                text['xml:space'] = next_text['xml:space']
              end
              next_r.remove
            else
              break
            end
          end

          if match = text.content.match(SECTION_TEXT)
            type = match[1]
            section_name = match[2]
            @sections[section_name] ||= {}
            @sections[section_name][type.downcase.to_sym] = p
          end

          text.content.scan(VAR_PATTERN).each do |match|
            if m = match[0].match(/(if|end)block ?(.*)/)
              if m[1] == 'if'
                if !last_ifblock.nil?
                  raise "Unmatched ifblock (#{last_ifblock})"
                end
                last_ifblock = block_id = m[2]
              elsif last_ifblock.nil?
                raise "Unmatched endblock"
              else
                block_id = last_ifblock
                last_ifblock = nil
              end
              (@label_nodes["ifblock #{block_id}"] ||= []) << r
            elsif m = match[0].match(/([\$0-9a-zA-Z_\-\.\[\]]+\[.*\])/)
              pieces = m[0].split('.')
              pieces.each_with_index do |piece, i|
                if piece.match /\[[^\[]*\]$/
                  (@label_nodes[pieces[0..i].join('.')] ||= []) << r
                end
              end
            end
          end

          cur_child += 1
        end
      end

      # Must be processed from inner to outer, so sort in desc. length
      @label_nodes.sort { |x, y| y[0].length - x[0].length }.each do |label, nodes|
        nodes.uniq!
        # Find lowest common ancestor
        if nodes.length == 1
          # Default to finding the cell if it's in a table
          # or the last paragraph if it isn't
          node = nodes[0]
          first_p = nil
          first_tc = nil
          while first_tc == nil && node.name != 'body'
            if node.name == 'tc'
              first_tc ||= node
            elsif node.name == 'p'
              first_p ||= node
            end
            node = node.parent
          end
          if root = first_tc || first_p
            @label_root[label] = root
            root['templateLabel'] = label
            @node_label[root] = label
          end
        else
          node_paths = nodes.map do |node|
            path = []
            cur_node = node.parent
            while cur_node.name != 'body'
              path << cur_node
              cur_node = cur_node.parent
            end
            path.reverse
          end

          low = 0
          high = node_paths.map { |p| p.length }.min - 1
          while low <= high
            mid = (low + high + 1) / 2
            if node_paths.all? { |path| path[mid] == node_paths[0][mid] }
              low = mid
              break if low == high
            else
              high = mid - 1
            end
          end

          if low == high && node_paths[0][low].name != 'tbl'
            node = node_paths[0][low]
            if @node_label[node].nil?
              @label_root[label] = node
              node['templateLabel'] = label
              @node_label[node] = label
            end
          else
            # Create a fake node containing all these paragraphs/rows (and everything in between)
            marked = Hash[node_paths.map { |p| [p[high + 1], true] }]
            body = high == -1 ? node_paths[0][0].parent : node_paths[0][low]
            first_child = -1
            last_child = -1
            body.children.each_with_index do |c, i|
              if marked[c]
                if first_child == -1
                  first_child = i
                end
                last_child = i
              end
            end

            after = body.children[last_child + 1]

            group_node = Nokogiri::XML::Node.new('templateGroup', body.document)
            body.children[first_child..last_child].each do |n|
              n.remove
              group_node.add_child n
            end

            if after.nil?
              body.children.after group_node
            else
              after.before group_node
            end

            group_node['templateLabel'] = label
            @node_label[group_node] = label
            @label_root[label] = group_node
          end
        end
      end
    end

    def top_level(node)
      while node.parent.name != 'body'
        node = node.parent
      end
      node
    end

    def replace_sections(root)
      if section = @options[:extract_section]
        if section = @sections[section]
          # Might be enclosed in templateGroups now.
          sec_start = top_level(section[:start])
          sec_end = top_level(section[:end])

          take = false
          root.children.each do |child|
            if child == sec_start
              take = true
            end
            if !take
              child.remove
            end
            if child == sec_end
              take = false
            end
          end
        end
      elsif @options[:replace_sections]
        section_starts = {}
        @options[:replace_sections].each do |section, _|
          if @sections[section]
            section_starts[@sections[section][:start]] = section
          else
            puts "Warning: section '#{section}'' not found."
          end
        end

        cur_section = nil
        root.children.each do |child|
          if section_starts[child]
            cur_section = section_starts[child]
            ins_doc = @options[:replace_sections][cur_section]
            prev_node = nil
            ins_doc.at_xpath('//w:document//w:body').children.each do |node|
              if node.name == 'p' || node.name == 'tbl'
                prev_node = node.dup
                child.add_previous_sibling(prev_node)
              elsif node.name == 'sectPr'
                # Insert this into a paragraph.
                if prev_node.nil? || prev_node.name != 'p'
                  prev_node = Nokogiri::XML::Node.new('w:p', root)
                  child.add_previous_sibling(prev_node)
                end
                pPr = prev_node.at_xpath('w:pPr')
                pPr ||= Nokogiri::XML::Node.new('w:pPr', root)
                pPr.add_child(node.dup)
                prev_node.children.before(pPr)
              end
            end
          end

          if cur_section
            child.remove
          end

          if cur_section && child == @sections[cur_section][:end]
            cur_section = nil
          end
        end
      end
    end

    def add_placeholder(str)
      m = str.match(VAR_PATTERN)
      return "<%= (#{m[1]}) || '#{@options[:placeholder]}' %>"
    end

    def process_indices(str)
      str.split(/ /).map { |s|
        s.gsub!(/(?<=[a-zA-Z\]}])\.(?=[a-zA-Z])/, '.andand.')

        cur_array = -1
        s.gsub(/\[[^\[]*\]/) do
          cur_array += 1
          "[$index[#{cur_array}]]"
        end
      }.join(' ')
    end

    def replace_runs(node, data, data_index = [])
      if node.name == 'p'
        cur_child = 0
        node.xpath('.//w:r').each do |r|
          text = find_child(r, 't')
          if !text.nil?
            $index = data_index
            content = text.content.gsub(VAR_PATTERN) do |match|
              add_placeholder(process_indices(match))
            end
            text.content = @erb_binding.render(content)
            wspace_regex = /[\t\n]|(^\s)|(\s$)|(\s\s)/
            if text.content.match(wspace_regex)
              text['xml:space'] = 'preserve'
            end
            # Extract checkbox symbols; put them into w:sym
            check_regex = /([#{CheckboxValue::CHECKED_CHAR}#{CheckboxValue::UNCHECKED_CHAR}])/
            if text.content[check_regex]
              cur = text
              text.content.split(check_regex).each do |p|
                if p[check_regex]
                  sym_node = Nokogiri::XML::Node.new 'w:sym', node
                  sym_node['w:font'] = 'Wingdings'
                  sym_node['w:char'] = p.ord.to_s(16)
                  cur = cur.add_next_sibling(sym_node)
                elsif !p.empty?
                  t_node = text.clone
                  t_node.content = p
                  if p[wspace_regex]
                    t_node['xml:space'] = 'preserve'
                  end
                  cur = cur.add_next_sibling(t_node)
                end
              end
              text.remove
            end
          end
        end
      else
        prev_child = nil
        node.children.each do |child|
          if label = child['templateLabel']
            $index = data_index

            if m = label.match(/^ifblock (.*)$/)
              dat = @erb_binding.eval(process_indices(m[1]))
              if dat
                dat = [:ifblock]
              else
                dat = nil
              end
            else
              dat = @erb_binding.eval(process_indices(label.gsub(/\[[^\[]*\]$/, '')))
            end

            if dat.nil? || dat.length == 0
              child.remove
            else
              master_copy = child.clone
              child.remove

              next_child = prev_child
              for i in 0..dat.length-1
                if next_child.nil?
                  next_child = node.before master_copy.clone
                else
                  next_child = next_child.add_next_sibling(master_copy.clone)
                end
                replace_runs(next_child, data, data_index + (dat[i] == :ifblock ? [] : [i]))
                prev_child = next_child
                if child.name == 'templateGroup'
                  orig_child = next_child
                  orig_child.children.each do |subchild|
                    next_child = next_child.add_next_sibling subchild
                    prev_child = next_child
                  end
                  orig_child.remove
                end
              end
            end
          else
            replace_runs(child, data, data_index)
            prev_child = child
          end
        end
      end
    end

    def remove_empty(node)
      if node.name == 'p'
        node.children.each do |child|
          if child.name == 'r'
            t = find_child(child, 't')
            if t && t.content[DELETE_TEXT]
              if node.parent.name == 'tc' && node.parent.xpath('w:p').length == 1
                node.children.remove
              else
                node.remove
              end
              return
            end
          end
        end
      else
        if node.name == 'tbl'
          node.xpath('w:tr').each do |tr|
            if tc = find_child(tr, 'tc')
              if t = tc.at_xpath('.//w:t')
                if t.content[DELETE_TEXT]
                  tr.remove
                end
              end
            end
          end

          if find_child(node, 'tr').nil?
            node.remove
            return
          end
        end

        node.children.each do |child|
          remove_empty(child)
        end
      end
    end

    # Trim trailing empty paragraphs.
    def trim_empty_paragraphs(node)
      node.children.reverse_each do |child|
        if child.name == 'sectPr'
          last_sect = child
        elsif child.name == 'p'
          empty = true
          child.children.each do |pchild|
            if pchild.name == 'pPr'
              # ignore
            elsif pchild.name == 'r'
              pchild.children.each do |rchild|
                if rchild.name == 't'
                  if !rchild.content.empty?
                    empty = false
                  end
                elsif rchild.name != 'rPr'
                  empty = false
                end
              end
            else
              empty = false
            end
          end

          if pPr = child.at_xpath('w:pPr')
            if sect = pPr.at_xpath('w:sectPr')
              # This is the true "last section". Move it to the end of the body.
              if last_sect
                last_sect.remove
              end
              pPr.remove
              node.children.after(sect)
              last_sect = sect
              if empty
                child.remove
              end
              break
            end
          end

          if empty
            child.remove
          else
            break
          end
        else
          break
        end
      end
    end
  end
end
