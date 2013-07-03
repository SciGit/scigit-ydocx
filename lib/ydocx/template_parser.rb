#!/usr/bin/env ruby
# encoding: utf-8

require 'nokogiri'

module YDocx
  class TemplateParser
    def initialize(doc)
      @label_nodes = {}
      @label_root = {}
      @node_label = {}
      @doc = doc
    end

    def replace(data)
      doc = Nokogiri::XML.parse(@doc)
      group_values(doc.at_xpath('//w:document//w:body'), data)
      replace_runs(doc.at_xpath('//w:document//w:body'), data)
      return doc
    end

    def lookup(str, data, data_index = nil)
      parts = str.split('.')
      if parts.length == 1
        return data[parts[0]]
      else
        # only handle 2 parts for now
        if data[parts[0]].is_a?(Array) && !data_index.nil?
          return data[parts[0]][data_index] && data[parts[0]][data_index][parts[1]]
        else
          return nil
        end
      end
    end

    def stringify(obj)
      if obj.nil?
        return 'N/A'
      elsif obj.is_a? Hash
        # checkbox
        true_keys = obj.select { |k,v| v }.keys
        return true_keys.join(', ')
      else
        return obj.to_s
      end
    end

    def group_values(node, data)
      node.xpath('.//w:r').each do |run|
        run.children.each do |t|
          if t.name == 't'
            t.content.scan(/%([0-9a-zA-Z_\-\.\[\]]+)%/).each do |match|
              parts = match[0].split('.')
              if parts.length > 1
                (@label_nodes[parts[0]] ||= []) << run
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
          high = node_paths.map { |p| p.length }.max - 1
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
              @node_label[node] = label
            end
          end
        end
      end
    end

    def replace_runs(node, data, data_index = nil)
      if node.name == 'p'
        cur_child = 0
        while cur_child < node.children.length
          r = node.children[cur_child]
          text = r.at_xpath('w:t')
          if text.nil?
            cur_child += 1
            next
          end

          prop = r.at_xpath('w:rPr')
          prop = prop ? prop.to_s : ''
          while cur_child + 1 < node.children.length
            next_r = node.children[cur_child + 1]
            next_prop = next_r.at_xpath('w:rPr')
            next_prop = next_prop ? next_prop.to_s : ''
            if next_r.name.start_with?('bookmark')
              next_r.remove # These are inserted randomly, can't really tell what they do
            elsif next_prop == prop && (next_text = next_r.at_xpath('w:t'))
              text.content += next_text.content
              next_r.remove
            else
              break
            end
          end

          last_index = 0
          processed = ''
          text.content.to_enum(:scan, /%([0-9a-zA-Z_\-\.\[\]]+)%/).each do
            match = Regexp.last_match
            index = match.pre_match.length
            if index > last_index
              processed += text.content[last_index..index-1]
            end
            processed += stringify(lookup(match[1], data, data_index))
            last_index = index + match[0].length
          end
          processed += text.content[last_index..-1]
          text.content = processed
          cur_child += 1
        end
        return
      end

      node.children.each do |child|
        if label = @node_label[child]
          if data[label].nil? || data[label].length == 0
            child.remove
          else
            next_child = child
            for i in 1..data[label].length-1
              next_child = next_child.add_next_sibling(child.clone)
              replace_runs(next_child, data, i)
            end
            replace_runs(child, data, 0)
          end
        else
          replace_runs(child, data, data_index)
        end
      end
    end
  end
end