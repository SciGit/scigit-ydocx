#!/usr/bin/env ruby
# encoding: utf-8

require 'ydocx'
require 'diff/lcs'

module YDocx
  class Differ
    def self.chunk_similarity(chunk1, chunk2)
      text1 = chunk1.map { |c| c.join('') }
      text2 = chunk2.map { |c| c.join('') }
      # Find the LCS between the runs
      lcs = Diff::LCS.LCS(text1, text2)
      lcs_len = 0
      lcs.each do |run|
        lcs_len += run.length
      end
      tlen = 0
      (text1 + text2).each do |run|
        tlen += run.length
      end
      if tlen == 0
        return 1
      else
        return 2.0 * lcs_len / tlen
      end
    end
    # Return the % similarity between two blocks (paragraphs, tables)
    def self.block_similarity(block1, block2)
      if block1.class != block2.class
        return 0
      end
      if block1.is_a? Paragraph
        chunk_similarity(block1.get_chunks, block2.get_chunks)
      elsif block1.is_a? Table
        return 1
      else
        return 0
      end
    end
    def self.get_chunks(paragraphs)
      chunks = []
      paragraphs.each_with_index do |p, i|
        if i > 0
          chunks << [Run.new("\r", Style.new)]
        end
        chunks += p.get_chunks
      end
      chunks
    end
    def self.diff(doc1, doc2)
      blocks1 = doc1.contents.blocks
      blocks2 = doc2.contents.blocks
      # Do an n^2 LCS diff on the blocks.
      n = blocks1.length
      m = blocks2.length
      lcs = Array.new(n+1) { Array.new(m+1, 0) }
      action = Array.new(n+1) { Array.new(m+1, -1) }
      blocks1.reverse.each_with_index do |a, ii|
        blocks2.reverse.each_with_index do |b, jj|
          if n*m > 1000
            sim = (a == b ? 1 : 0)
          else
            sim = block_similarity(a, b)
          end
          i = n-1-ii
          j = m-1-jj
          lcs[i][j] = lcs[i+1][j]
          action[i][j] = 0
          if lcs[i][j+1] > lcs[i][j]
            lcs[i][j] = lcs[i][j+1]
            action[i][j] = 1
          end
          if sim > 0.5 && lcs[i+1][j+1] + sim > lcs[i][j]
            lcs[i][j] = lcs[i+1][j+1] + sim
            action[i][j] = 2
          end
        end
      end
      i = 0
      j = 0
      lblocks = []
      rblocks = []
      diff_blocks = []
      while i < n || j < m        
        if j == m || action[i][j] == 0
          lblocks << blocks1[i]
          i += 1
        elsif i == n || action[i][j] == 1
          rblocks << blocks2[j]
          j += 1
        else
          unless lblocks.empty? && rblocks.empty?
            if lblocks.empty? || rblocks.empty? ||
               chunk_similarity(get_chunks(lblocks), get_chunks(rblocks)) > 0.5
              diff_blocks << [lblocks.dup, rblocks.dup]
            else
              puts lblocks.to_s
              puts rblocks.to_s
              diff_blocks << [lblocks.dup, []]
              diff_blocks << [[], rblocks.dup]
            end
            lblocks = []
            rblocks = []
          end
          diff_blocks << [[blocks1[i]], [blocks2[j]]]
          i += 1
          j += 1
        end
      end
      unless lblocks.empty? && rblocks.empty?
        if lblocks.empty? || rblocks.empty? ||
           chunk_similarity(get_chunks(lblocks), get_chunks(rblocks)) > 0.5
          diff_blocks << [lblocks.dup, rblocks.dup]
        else
          diff_blocks << [lblocks.dup, []]
          diff_blocks << [[], rblocks.dup]
        end
      end
      
      table = Table.new
      diff_blocks.each do |block|
        c1 = Cell.new
        c1.blocks = block[0]
        c2 = Cell.new
        c2.blocks = block[1]
        if block[0].empty?
          c2.class = 'add'
        elsif block[1].empty?
          c1.class = 'delete'
        elsif block[0] != block[1]
          c1.class = c2.class = 'modify'
        end
        table.cells << [c1, c2]
      end
      
      [doc1, doc2].each do |doc|
        if !doc.images.empty?
          doc.create_files
        end
      end
      
      html_doc = ParsedDocument.new
      html_doc.blocks << table
      builder = Builder.new(html_doc)
      builder.title = 'Diff Results'
      builder.style = true
      builder.build_html
    end
  end
end