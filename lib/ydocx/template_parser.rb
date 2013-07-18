#!/usr/bin/env ruby
# encoding: utf-8

require 'nokogiri'

module YDocx
  class TemplateParser
    OPTION_DEFAULTS = {
      :placeholder => 'None or N/A'
    }

    VAR_PATTERN = /%([0-9a-zA-Z_\-\.\[\]]+)(|[^%\+]*)?(\+[^%]*)?%/

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
        else
          store[field[:id]] = field
        end
      end
    end

    def unwrap(data)
      if data.is_a?(Hash)
        if data['value']
          return unwrap(data['value'])
        else
          data.each do |k, v|
            data[k] = unwrap(v)
          end
        end
      elsif data.is_a?(Array)
        return data.map { |d| unwrap(d) }
      end

      return data
    end

    def replace(data)
      data = unwrap(data)

      doc = Nokogiri::XML.parse(@doc)
      merge_runs(doc)
      root = doc.at_xpath('//w:document//w:body')
      group_values(root, data)
      replace_runs(root, data)
      return doc
    end

    def rec_lookup(parts, fields, data, data_index = [])
      if parts.length == 1
        if fields[:type] == 'checkbox'
          return [data && data[parts[0]], fields, parts[0]]
        elsif fields[parts[0]].nil?
          #p parts
          return nil
        else
          return [data && data[parts[0]], fields[parts[0]]]
        end
      else
        if fields[parts[0]].nil?
          #p parts
          return nil
        elsif fields[parts[0]][:type].nil?
          data = data && data[parts[0]]
          if data.is_a?(Hash) && !data['value'].nil?
            data = data['value']
          end
          return rec_lookup(parts[1..-1], fields[parts[0]], data && data[data_index[0]], data_index[1..-1])
        else
          return rec_lookup(parts[1..-1], fields[parts[0]], data && data[parts[0]], data_index)
        end
      end
    end

    def lookup(str, data, data_index = [])
      parts = str.split('.')
      rec_lookup(parts, @fields, data, data_index)
    end

    def stringify(obj, default = nil, after = nil)
      if obj.nil?
        return default || @options[:placeholder]
      end

      val = obj[0]
      field = obj[1]
      if field[:type] == 'checkbox'
        return (val ? '&#9632; ' : '&#9633;')  + ' ' + field[:options][obj[2]]
      end

      if val.nil?
        return default || @options[:placeholder]
      end

      if field[:type] == 'radio'
        return field[:options][val]
      elsif field[:type] == 'currency'
        return '$' + val
      else
        return val.to_s
      end
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
              parts = match[0].split('.')
              if parts.length > 1
                field = @fields
                parts.each_with_index do |p, i|
                  field = field[p]
                  if field.nil?
                    break
                  end
                  if field[:id].nil?
                    (@label_nodes[parts[0..i].join('.')] ||= []) << run
                  end
                end
              end
            end
          end
        end
      end

      @label_nodes.each do |label, nodes|
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

    def replace_runs(node, data, data_index = [])
      # process if statements
      if node.name == 'tbl'
        if tr = node.at_xpath('w:tr')
          if tc = tr.at_xpath('w:tc')
            if t = tc.at_xpath('.//w:t')
              value = ''
              t.content.scan(/%(show)?if ([0-9a-zA-Z_\-\.\[\]]+)%/).each do |match|
                value = lookup(match[1], data, data_index)
                value = value && value.first
                if value.nil? || value.empty? || value.downcase == 'no'
                  node.remove
                  return
                end
              end
              t.inner_html = t.content.gsub(/%if ([0-9a-zA-Z_\-\.\[\]]+)%/, '')
              t.inner_html = t.content.gsub(/%showif ([0-9a-zA-Z_\-\.\[\]]+)%/, value)
            end
          end
        end
      end

      if node.name == 'p'
        if t = node.at_xpath('.//w:t')
          value = ''
          t.content.scan(/%(show)?if ([0-9a-zA-Z_\-\.\[\]]+)%/).each do |match|
            value = lookup(match[1], data, data_index)
            value = value && value.first.to_s
            if value.nil? || value.empty? || value.downcase == 'no' || value.downcase == 'false'
              node.remove
              return
            end
          end
          t.inner_html = t.content.gsub(/%if ([0-9a-zA-Z_\-\.\[\]]+)%/, '')
          t.inner_html = t.content.gsub(/%showif ([0-9a-zA-Z_\-\.\[\]]+)%/, value)
        end

        cur_child = 0
        while cur_child < node.children.length
          r = node.children[cur_child]
          text = r.at_xpath('w:t')
          if text.nil?
            cur_child += 1
            next
          end

          last_index = 0
          processed = ''
          text.content.to_enum(:scan, VAR_PATTERN).each do
            match = Regexp.last_match
            index = match.pre_match.length
            if index > last_index
              processed += text.content[last_index..index-1]
            end
            processed += stringify(lookup(match[1], data, data_index), match[2] && match[2][1..-1], match[3] && match[3][1..-1])
            last_index = index + match[0].length
          end
          processed += text.content[last_index..-1]
          text.inner_html = processed
          cur_child += 1
        end
        return
      end

      prev_child = nil
      node.children.each do |child|
        if label = child['templateLabel']
          dat = lookup(label, data, data_index)
          dat = dat && dat.first

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
                orig_child = remove
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
end
