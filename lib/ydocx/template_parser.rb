#!/usr/bin/env ruby
# encoding: utf-8

require 'nokogiri'
require 'erb'
require 'ostruct'
require 'andand'

def to_bool(value)
  if value.nil?
    return false
  end
  value = value.to_s.downcase
  return !value.empty? && value != 'no' && value != 'false' && value != 'null'
end

module YDocx
  DELETE_TEXT = '###DELETE_ME###'

  class ErbBinding < OpenStruct
    def render(template)
      ERB.new(template).result(binding)
    end

    def eval(x)
      binding.eval(x)
    end

    def showif(x)
      return to_bool(x) ? '' : DELETE_TEXT
    end

    def selfif(x)
      return (to_bool(x) && x) || DELETE_TEXT
    end

    def append(x, y, placeholder = '')
      return x && x.to_s + y || placeholder
    end

    def date(x)
      if x.to_i
        return Time.at(x.to_i).strftime('%Y/%-m/%-d')
      else
        return nil
      end
    end
  end

  class TemplateParser
    OPTION_DEFAULTS = {
      :placeholder => 'None or N/A'
    }

    VAR_PATTERN = /<%=(([^%]|%(?!>))*)%>/

    def initialize(doc, fields, options)
      @label_nodes = {}
      @label_root = {}
      @node_label = {}
      @doc = doc
      @fields = {}
      @options = options.merge(OPTION_DEFAULTS)
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
            vals[k] = (to_bool(data.andand[k]) ? '&#9632; ' : '&#9633;')  + ' ' + v
          end
          return OpenStruct.new(vals)
        when 'radio'
          return fields[:options][data]
        when 'currency'
          return data && !data.empty? && '$' + data
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
      merge_runs(doc)
      root = doc.at_xpath('//w:document//w:body')
      group_values(root, data)
      replace_runs(root, data)
      remove_empty(root)
      return doc
    end

    def merge_runs(node)
      node.xpath('.//w:p').each do |p|
        cur_child = 0
        while cur_child < p.children.length
          r = p.children[cur_child]
          text = r.at_xpath('w:t')
          if text.nil?
            cur_child += 1
            next
          end

          prop = r.at_xpath('w:rPr')
          prop = prop ? prop.to_s : ''
          while cur_child + 1 < p.children.length
            next_r = p.children[cur_child + 1]
            next_prop = next_r.at_xpath('w:rPr')
            next_prop = next_prop ? next_prop.to_s : ''
            if next_r.name.start_with?('bookmark') || next_r.name == 'proofErr'
              next_r.remove # These are inserted randomly, can't really tell what they do
            elsif next_prop == prop && (next_text = next_r.at_xpath('w:t'))
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

          cur_child += 1
        end
      end
    end

    def group_values(node, data)
      node.xpath('.//w:r').each do |run|
        run.children.each do |t|
          if t.name == 't'
            t.content.scan(VAR_PATTERN).each do |match|
              if m = match[0].match(/([\$0-9a-zA-Z_\-\.\[.*\]]+\[.*\])/)
                pieces = m[0].split('.')
                pieces.each_with_index do |piece, i|
                  if piece.match /\[[^\[]*\]$/
                    (@label_nodes[pieces[0..i].join('.')] ||= []) << run
                  end
                end
              end
            end
          end
        end
      end

      # Must be processed from inner to outer, so sort in desc. length
      @label_nodes.sort { |x, y| y[0].length - x[0].length }.each do |label, nodes|
        # Find lowest common ancestor
        if nodes.length == 1
          # Default to finding the row if it's in a table
          # or the last paragraph if it isn't
          node = nodes[0]
          first_p = nil
          first_tr = nil
          while first_tr == nil && node.name != 'body'
            if node.name == 'tr'
              first_tr ||= node
            elsif node.name == 'p'
              first_p ||= node
            end
            node = node.parent
          end
          if root = first_tr || first_p
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

          if low == high
            node = node_paths[0][low]
            if @node_label[node].nil?
              @label_root[label] = node
              node['templateLabel'] = label
              @node_label[node] = label
            end
          else
            # Create a fake node containing all these paragraphs (and everything in between)
            marked = Hash[node_paths.map { |p| [p[0], true] }]
            body = node_paths[0][0].parent
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

    def add_placeholder(str)
      m = str.match(VAR_PATTERN)
      return "<%= (#{m[1]}) || '#{@options[:placeholder]}' %>"
    end

    def process_indices(str)
      cur_array = -1
      str = str.gsub('.', '.andand.')
      str.gsub(/\[[^\[]*\]/) do
        cur_array += 1
        "[$index[#{cur_array}]]"
      end
    end

    def replace_runs(node, data, data_index = [])
      if node.name == 'p'
        cur_child = 0
        node.children.each do |r|
          text = r.at_xpath('w:t')
          if !text.nil?
            $index = data_index
            content = text.content.gsub(VAR_PATTERN) do |match|
              add_placeholder(process_indices(match))
            end
            text.inner_html = @erb_binding.render(content)
          end
        end
      else
        prev_child = nil
        node.children.each do |child|
          if label = child['templateLabel']
            $index = data_index
            dat = @erb_binding.eval(process_indices(label[0..-3]))

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
                replace_runs(next_child, data, data_index + [i])
                if node.name == 'templateGroup'
                  orig_child = next_child
                  orig_child.children.each do |subchild|
                    next_child = next_child.add_next_sibling subchild
                  end
                  orig_child.remove
                end
              end
            end
          else
            replace_runs(child, data, data_index)
          end
          prev_child = child
        end        
      end
    end

    def remove_empty(node)
      if node.name == 'p'
        if t = node.at_xpath('.//w:t')
          if t.content[DELETE_TEXT]
            if node.parent.name == 'tc' && node.parent.xpath('w:p').length == 1
              node.children.remove
            else
              node.remove
            end
            return
          end
        end
      else
        if node.name == 'tbl'
          node.xpath('w:tr').each do |tr|
            if tc = tr.at_xpath('w:tc')
              if t = tc.at_xpath('.//w:t')
                if t.content[DELETE_TEXT]
                  tr.remove
                end
              end
            end
          end

          if node.at_xpath('w:tr').nil?
            node.remove
            return
          end
        end

        node.children.each do |child|
          remove_empty(child)
        end
      end
    end
  end
end
