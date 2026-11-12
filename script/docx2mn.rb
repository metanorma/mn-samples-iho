#!/usr/bin/env ruby
# frozen_string_literal: true

# DOCX to Metanorma (IHO flavor) AsciiDoc converter built on uniword.
#
# Reads a Word document through the uniword object model (no direct XML
# scraping of document.xml) and emits a Metanorma document skeleton
# (document.adoc + sections/*.adoc + images/) in the shape used by the
# mn-samples-iho repository (see sources/s102 for the reference shape).
#
# Handles: ordered inline traversal (element_order), headings and appendix
# headings with obligation, bullet/numbered lists from numbering
# definitions, tables with colspan and complex cells, figures with captions,
# footnotes/endnotes, ASN.1 source blocks, bibliography sections,
# bookmark->anchor mapping for internal references (REF fields and internal
# hyperlinks), external hyperlinks, and media extraction.
#
# Usage (run from the uniword checkout so its bundle is used):
#   cd ~/src/mn/uniword
#   bundle exec ruby /path/to/mn-samples-iho/script/docx2mn.rb \
#     input.docx OUT_DIR --kind main|annex-a|annex-b

require "fileutils"
require "optparse"
require "uri"
require "zip"

# Load uniword. When run outside the uniword bundle, pull in its lib
# directory from the sibling checkout. (Do not guard with
# `defined?(Uniword)`: Bundler evaluates uniword.gemspec, which partially
# defines Uniword before lib/uniword.rb has run.)
sibling = File.expand_path("../../uniword/lib", __dir__)
unless Gem::Specification.find_all_by_name("uniword").any? ||
       File.directory?(sibling)
  warn "uniword gem not found; run from the uniword checkout (bundle exec)"
  exit 1
end
$LOAD_PATH.unshift(sibling) if File.directory?(sibling) &&
                                Gem::Specification.find_all_by_name("uniword").empty?
require "uniword"

module DocxToMn
  W = Uniword::Wordprocessingml

  DOC_CONFIG = {
    "main" => {
      header_title: "Maritime Limits and Boundaries Product Specification",
      annex_attr: nil,
    },
    "annex-a" => {
      header_title: "Maritime Limits and Boundaries Product Specification",
      annex_attr: "Annex A – Data Classification and Encoding Guide",
    },
    "annex-b" => {
      header_title: "Maritime Limits and Boundaries Product Specification",
      annex_attr: "Annex B – Data Product Format (Encoding)",
    },
  }.freeze

  class Converter
    Block = Struct.new(
      :type, :para, :paras, :table, :level, :style, :num_id, :ilvl, :fmt,
      :caption, :caption_kind, :anchor, :anchors, :images, :obligation,
      :is_appendix, :biblio, keyword_init: true
    )

    Segment = Struct.new(:text, :raw, :bold, :italic, :sup, :sub,
                         keyword_init: true)

    PARAGRAPH_CHILDREN = {
      "r" => :runs, "hyperlink" => :hyperlinks,
      "bookmarkStart" => :bookmark_starts, "bookmarkEnd" => :bookmark_ends,
      "fldChar" => :field_chars, "instrText" => :instr_text,
      "fldSimple" => :simple_fields, "oMath" => :o_maths,
      "oMathPara" => :o_math_paras, "sdt" => :sdts,
    }.freeze

    RUN_CHILDREN = {
      "t" => :text, "br" => :break, "tab" => :tab, "drawing" => :drawings,
      "pict" => :pictures, "noBreakHyphen" => :no_break_hyphen, "sym" => :sym,
    }.freeze

    attr_reader :warnings

    def initialize(docx_path, out_dir, kind)
      @docx_path = docx_path
      @out_dir = out_dir
      @kind = kind
      @doc = Uniword::DocumentFactory.from_file(docx_path)
      @zip = Zip::File.open(docx_path)
      @warnings = []
      @used_ids = {}
      @anchor_for_bookmark = {}
      @blocks = []
      @rels = build_rels
      @numbering_cache = {}
      @images = []
      @fig_seq = 0
      @tbl_seq = 0
      @in_toc_zone = false
    end

    def run
      classify
      postprocess
      assign_anchors
      write_output
    end

    private

    def warn_once(msg)
      @warnings << msg unless @warnings.include?(msg)
    end

    # ---------------------------------------------------------------- rels

    def build_rels
      Array(@doc.document_rels&.relationships).each_with_object({}) do |r, h|
        h[r.id] = { target: r.target.to_s, mode: r.target_mode.to_s }
      end
    end

    # ----------------------------------------------------------- numbering

    def numbering_levels(num_id)
      @numbering_cache[num_id] ||=
        begin
          d = @doc.numbering_configuration.get_definition_for_num_id(num_id)
          d ? d.levels.to_a : []
        end
    end

    def list_format(num_id, ilvl)
      numbering_levels(num_id)[ilvl.to_i]&.format_value || "decimal"
    end

    # ---------------------------------------------------- phase 1: classify

    def style_name(para)
      Array(para.properties&.style).first&.value.to_s
    end

    def numpr(para)
      np = para.properties&.numbering_properties
      return nil if np.nil? || np.num_id.nil?

      id = np.num_id.value.to_i
      return nil if id.zero?

      [id, np.ilvl&.value.to_i]
    end

    def paragraph_drawings(para)
      Array(para.runs).flat_map { |r| Array(r.drawings) }
    end

    BODY_CHILDREN = {
      "p" => :paragraphs, "tbl" => :tables, "sdt" => :structured_document_tags,
    }.freeze

    # Body#elements in uniword returns paragraphs before tables regardless of
    # document order; reconstruct the true order from element_order.
    def body_elements_in_order
      ordered = []
      each_named(@doc.body, BODY_CHILDREN) do |_name, item|
        ordered << item
      end
      ordered
    end

    def classify
      elements = expand_sdt(body_elements_in_order)
      elements.each_with_index do |el, idx|
        case el
        when W::Paragraph
          classify_paragraph(el, elements, idx)
        when W::Table
          @blocks << Block.new(type: :table, table: el)
        end
      end
    end

    def expand_sdt(elements)
      elements.flat_map do |el|
        if W::StructuredDocumentTag === el && el.content
          content = el.content
          expand_sdt(Array(content.paragraphs) + Array(content.tables))
        else
          [el]
        end
      end
    end

    def classify_paragraph(para, elements, idx)
      style = style_name(para)
      text = para.text.strip
      num = numpr(para)

      blk = Block.new(type: :body, para: para, style: style, anchors: [])

      case style.downcase
      when "toctitle"
        if text.downcase == "document control"
          blk.type = :preface
        else
          blk.type = :skip
          @in_toc_zone = true if text =~ /table of contents/i
        end
      when /\Atoctitle\d\z/, /\Atoc\d\z/, "zzcover", "title", /\Atitle\d\z/
        blk.type = :skip
      when "subtitle"
        blk.type = :skip
        attach_obligation(text)
      when /\Aheading(\d)\z/
        blk.type = :heading
        blk.level = Regexp.last_match(1).to_i + 1
        @in_toc_zone = false
      when /\Aapph(-[a-z]\d?)?\z/
        classify_apph(blk, num, elements, idx)
        @in_toc_zone = false
      when "figurecaption", "caption"
        if text.empty? && paragraph_drawings(para).any?
          blk.type = :body # image inside a caption-styled paragraph: figure
        else
          blk.type = :caption
          blk.caption_kind = text.match?(/\Afigure/i) ? :figure : :table
        end
      when "asn1", "code"
        blk.type = :source_line
      when "notes"
        blk.type = text.empty? ? :skip : :note
      when "normreference", "bibliography1"
        blk.type = :bib
      else
        if text.match?(/\A\((normative|informative)\)\z/i)
          blk.type = :skip
          attach_obligation(text)
        else
          blk.type = classify_body(para, text, num)
        end
      end

      if blk.type == :list_item
        blk.num_id = num.first
        blk.ilvl = num.last
        blk.fmt = list_format(num.first, num.last)
      end
      @blocks << blk
    end

    def classify_body(para, text, num)
      if text.empty?
        return :body if paragraph_drawings(para).any?
        return :skip if Array(para.bookmark_starts).empty?

        return :body # bookmark-only paragraph: inline anchor holder
      end
      return :skip if text.downcase == "page intentionally left blank"
      return :note if text.match?(/\A\s*note\b\s*\d*\s*[-–—:.]/i)
      return :list_item if num

      :body
    end

    # AppH-* styles are appendix headings; depth comes from numbering ilvl:
    # 0 => appendix title (==), 1 => clause (===), 2 => sub-clause (====).
    # Without numbering: an appendix title when followed by a
    # "(Normative)"/"(Informative)" subtitle, otherwise a sub-heading of the
    # current appendix (main document "Test case for ..." entries).
    def classify_apph(blk, num, elements, idx)
      followed_by_subtitle = false
      nxt = elements[idx + 1]
      if W::Paragraph === nxt &&
         nxt.text.strip.match?(/\A\((normative|informative)\)\z/i)
        followed_by_subtitle = true
      end

      blk.type = :heading
      if num
        blk.level = num.last.zero? ? 2 : num.last + 2
        blk.is_appendix = blk.level == 2
      else
        blk.level = followed_by_subtitle ? 2 : 3
        blk.is_appendix = followed_by_subtitle
      end
    end

    def attach_obligation(text)
      m = text.match(/\((normative|informative)\)/i)
      return unless m

      @blocks.reverse_each.find { |b| b.type == :heading }&.obligation =
        m[1].downcase
    end

    # ----------------------------------------------- phase 2: postprocess

    def postprocess
      mark_inline_captions
      convert_figures
      normalize_lists
      merge_source_lines
      attach_captions
      mark_biblio_headings
      scan_references
    end

    # Word lists frequently start at ilvl > 0; AsciiDoc lists must start at
    # the outermost marker, so offset each consecutive run by its min level.
    def normalize_lists
      run = []
      flush = lambda do
        min = run.map(&:ilvl).min
        run.each { |b| b.ilvl -= min } if min&.positive?
        run.clear
      end
      @blocks.each do |b|
        if b.type == :list_item
          run << b
        else
          flush.call
        end
      end
      flush.call
    end

    # Only bookmarks actually referenced somewhere outside the TOC get inline
    # anchors; structural anchors (headings/tables/figures) are always kept.
    def scan_references
      @referenced = {}
      walk_paragraphs(@blocks) do |p, _block|
        Array(p.hyperlinks).each { |h| @referenced[h.anchor] = true if h.anchor }
        Array(p.simple_fields).each do |f|
          m = f.instr.to_s.match(/\A\s*REF\s+(\S+)/i)
          @referenced[m[1]] = true if m
        end
        Array(p.instr_text).each do |it|
          m = it.content.to_s.match(/\A\s*REF\s+(\S+)/i)
          @referenced[m[1]] = true if m
        end
      end
    end

    def walk_paragraphs(blocks)
      blocks.each do |b|
        case b.type
        when :table
          table_paragraphs(b.table).each { |p| yield p, b }
        when :source
          Array(b.paras).each { |p| yield p, b }
        when :skip, :preface
          next
        else
          yield b.para, b if b.para
        end
      end
    end

    # Body paragraphs that are plainly captions ("Table 3-2 – ...",
    # "Figure B-1 – ...").
    def mark_inline_captions
      @blocks.each_with_index do |b, i|
        next unless b.type == :body && !b.anchors.any?

        txt = b.para.text.strip
        if txt.match?(/\Atable\s+[\dA-Za-z.\-]+\s*[-–—]/i) &&
           @blocks[i + 1]&.type == :table
          b.type = :caption
          b.caption_kind = :table
        elsif txt.match?(/\Afigure\s+[\dA-Za-z.\-]+\s*[-–—]/i)
          b.type = :caption
          b.caption_kind = :figure
        end
      end
    end

    # Standalone image paragraphs become figure blocks.
    def convert_figures
      @blocks.map! do |b|
        next b unless b.type == :body && b.para.text.strip.empty? &&
                      paragraph_drawings(b.para).any?

        b.type = :figure
        b.images = paragraph_drawings(b.para)
        b
      end
    end

    def merge_source_lines
      merged = []
      buffer = []
      flush = lambda do
        unless buffer.empty?
          merged << Block.new(type: :source, paras: buffer.dup)
          buffer.clear
        end
      end
      @blocks.each do |b|
        if b.type == :source_line
          buffer << b.para
        else
          flush.call
          merged << b
        end
      end
      flush.call
      @blocks = merged
    end

    def attach_captions
      result = []
      pending = nil
      @blocks.each do |b|
        if b.type == :caption
          pending = b
          next
        end
        if pending
          result << pending unless attach_caption(pending, result, b)
          pending = nil
        end
        result << b
      end
      if pending
        result << pending unless attach_caption(pending, result, nil)
      end
      @blocks = result
    end

    def attach_caption(cap, result, next_block)
      if cap.caption_kind == :figure
        if next_block&.type == :figure
          next_block.caption = cap.para.text.strip
          return true
        end
        prev = result.reverse_each.find { |x| %i[figure body].include?(x.type) }
        if prev&.type == :figure
          prev.caption = cap.para.text.strip
          return true
        end
      elsif cap.caption_kind == :table && next_block&.type == :table
        next_block.caption = cap.para.text.strip
        return true
      end
      false
    end

    def mark_biblio_headings
      @blocks.each_with_index do |b, i|
        next unless b.type == :heading

        nxt = @blocks[i + 1]
        b.biblio = true if nxt&.type == :bib
        b.biblio ||= b.para.text.strip.match?(
          /\A(bibliography|normative references|references)\z/i
        )
      end
    end

    # -------------------------------------------------- phase 3: anchors

    def slugify(text)
      s = text.to_s.downcase.gsub(/[^a-z0-9]+/, "-").gsub(/\A-+|-+\z/, "")[0, 60]
      s.to_s.empty? ? "x" : s
    end

    def unique_id(base)
      id = base
      n = 1
      while @used_ids.key?(id)
        n += 1
        id = "#{base}-#{n}"
      end
      @used_ids[id] = true
      id
    end

    def assign_anchors
      @blocks.each do |b|
        case b.type
        when :heading
          title = b.para.text.strip
          prefix = b.is_appendix ? "annex" : "sec"
          b.anchor = unique_id("#{prefix}-#{slugify(title)}")
        when :figure
          cap = caption_body(b.caption, "figure")
          @fig_seq += 1
          b.anchor = unique_id(cap.empty? ? "fig-#{@fig_seq}" : "fig-#{slugify(cap)}")
        when :table
          cap = caption_body(b.caption, "table")
          @tbl_seq += 1
          b.anchor = unique_id(cap.empty? ? "tbl-#{@tbl_seq}" : "tbl-#{slugify(cap)}")
        end
      end
      collect_bookmarks
    end

    def caption_body(caption, kind)
      body = caption.to_s
      stripped = body.sub(/\A#{kind}\s*[\dA-Za-z.\-–]*\s*[-–—:]\s*/i, "")
      return stripped.strip if stripped != body && !stripped.strip.empty?

      # "Figure A-3 Shows the classes" (no dash): strip the label only when
      # the remainder reads like a title.
      m = body.match(/\A#{kind}\s+[A-Za-z]?-?[\d.]+(?:\.\d+)*\s+(.*)\z/mi)
      return body.strip if m.nil? || m[1].empty?

      rest = m[1].strip
      rest.match?(/\A[A-Z0-9"“]/) ? rest : body.strip
    end

    def table_paragraphs(tbl)
      tbl.rows.flat_map { |r| r.cells.flat_map { |c| Array(c.paragraphs) } }
    end

    # Map every bookmark to an anchor: to a structural anchor when the
    # bookmark lives in a heading/table/figure, else to an inline anchor at
    # its paragraph position. Guarantees every internal reference resolves.
    def collect_bookmarks
      @blocks.each do |b|
        paras = case b.type
                when :heading then [b.para]
                when :table then table_paragraphs(b.table)
                when :figure then [b.para]
                else []
                end
        paras.each do |p|
          Array(p.bookmark_starts).each do |bm|
            name = bm.name.to_s
            target = b.anchor if %i[heading table figure].include?(b.type)
            @anchor_for_bookmark[name] ||= target if target
          end
        end
      end

      @blocks.each do |b|
        next unless %i[body bib note caption source].include?(b.type)

        paras = b.type == :source ? Array(b.paras) : [b.para]
        paras.each do |p|
          Array(p.bookmark_starts).each do |bm|
            name = bm.name.to_s
            next if @anchor_for_bookmark.key?(name)
            next unless @referenced&.key?(name)

            id = unique_id("bm-#{slugify(name)}")
            @anchor_for_bookmark[name] = id
            (b.anchors ||= []) << id
          end
        end
      end
    end

    # -------------------------------------------------- ordered traversal

    # Yields [element_name, model_object] in document order, matching each
    # element_order entry against the parsed collection by position.
    def each_named(model, mapping)
      counters = Hash.new(0)
      Array(model.element_order).each do |node|
        next unless node.respond_to?(:name) && node.name

        attr = mapping[node.name]
        next if attr.nil?

        collection = model.public_send(attr)
        if collection.is_a?(Array)
          item = collection[counters[node.name]]
          counters[node.name] += 1
        else
          item = collection
        end
        yield node.name, item if item
      end
    end

    # -------------------------------------------------- inline rendering

    def on_off(wrapper)
      return false if wrapper.nil?

      v = wrapper.respond_to?(:value) ? wrapper.value : wrapper
      v != false && v != "false" && v != "0" && v != "off" && !v.nil?
    end

    def render_inline(para, cell: false)
      segs = []
      field = nil
      each_named(para, PARAGRAPH_CHILDREN) do |name, item|
        case name
        when "r"
          if field && field[:phase] == :instr
            nil
          elsif field && field[:phase] == :display
            field[:display] << plain_run_text(item)
          else
            segs.concat(run_segments(item, cell: cell))
          end
        when "hyperlink"
          segs.concat(Array(hyperlink_segments(item, cell: cell)))
        when "fldChar"
          field = field_transition(field, item)
          if field && field[:done]
            segs << render_field(field[:instr], field[:display])
            field = nil
          end
        when "instrText"
          field[:instr] << item.content if field && field[:phase] == :instr
        when "fldSimple"
          segs << simple_field(item)
        when "oMath", "oMathPara", "sdt"
          warn_once("inline math/SDT content skipped")
        end
      end
      join_segments(segs, cell: cell)
    end

    def field_transition(field, fld_char)
      type = fld_char.field_type.respond_to?(:value) ? fld_char.field_type&.value : fld_char.field_type
      case type
      when "begin" then { instr: +"", display: [], phase: :instr, done: false }
      when "separate" then field&.merge(phase: :display)
      when "end" then field&.merge(done: true)
      else field
      end
    end

    def render_field(instr, display)
      text = display.join.strip
      case instr.strip
      when /\A\s*REF\s+(\S+)/i
        xref(Regexp.last_match(1), text)
      when /\A\s*HYPERLINK\s+"?([^"\s]+)"?/i
        "link:#{Regexp.last_match(1)}[#{text}]"
      when /\A\s*(PAGEREF|PAGE|NUMPAGES|SEQ|NOTEREF)\b/i
        ""
      else
        text
      end
    end

    def simple_field(fld)
      text = fld.runs.map { |r| plain_run_text(r) }.join.strip
      render_field(fld.instr.to_s, [text])
    end

    def xref(bookmark_name, text)
      target = @anchor_for_bookmark[bookmark_name]
      unless target
        warn_once("unresolved internal reference to bookmark #{bookmark_name}")
        return text
      end
      text.empty? ? "<<#{target}>>" : "<<#{target},#{text}>>"
    end

    def hyperlink_segments(link, cell: false)
      segs = link.runs.flat_map { |r| run_segments(r, cell: cell) }
      text = join_segments(segs, cell: cell)
      return Segment.new(raw: xref(link.anchor, text.strip)) if link.anchor

      url = @rels[link.id]&.fetch(:target)
      if url.nil? || url.empty?
        warn_once("hyperlink #{link.id} without target; text kept")
        Segment.new(raw: text)
      elsif url =~ /\Ahttps?:\/\//
        Segment.new(raw: "link:#{url}[#{text}]")
      else
        warn_once("non-external hyperlink relationship #{link.id}")
        Segment.new(raw: text)
      end
    end

    def plain_run_text(run)
      parts = []
      each_named(run, RUN_CHILDREN) do |name, item|
        case name
        when "t" then parts << item.content.to_s
        when "br" then parts << " "
        when "tab" then parts << " "
        when "noBreakHyphen" then parts << "-"
        end
      end
      parts.join
    end

    def run_segments(run, cell: false)
      props = run.properties
      bold = on_off(props&.bold)
      italic = on_off(props&.italic)
      valign = props&.vertical_align&.value
      segs = []
      each_named(run, RUN_CHILDREN) do |name, item|
        seg = Segment.new(bold: bold, italic: italic,
                          sup: valign == "superscript",
                          sub: valign == "subscript")
        case name
        when "t"
          seg.text = item.content.to_s
          segs << seg
        when "br"
          segs << Segment.new(raw: "\n")
        when "tab"
          seg.text = " "
          segs << seg
        when "noBreakHyphen"
          seg.text = "-"
          segs << seg
        when "drawing"
          segs.concat(Array(drawing_segments(item)))
        when "pict"
          warn_once("legacy VML picture skipped")
        when "sym"
          warn_once("symbol character skipped")
        end
      end
      segs << Segment.new(raw: footnote_macro(run.footnote_reference)) if run.footnote_reference
      segs << Segment.new(raw: endnote_macro(run.endnote_reference)) if run.endnote_reference
      segs
    end

    def footnote_macro(ref)
      content = notes_text(@doc.footnotes, ref.id)
      content.empty? ? "" : "footnote:fn#{ref.id}[#{escape_macro_text(content)}]"
    end

    def endnote_macro(ref)
      content = notes_text(@doc.endnotes, ref.id)
      content.empty? ? "" : "footnote:en#{ref.id}[#{escape_macro_text(content)}]"
    end

    # [, ], \, * and _ inside an inline macro body terminate or corrupt the
    # macro parse (e.g. a footnote whose content starts with "*.txt*").
    def escape_macro_text(text)
      text.gsub(/([\[\]\\*_])/, "\\\\\\1")
    end

    def notes_text(container, id)
      entries = container.respond_to?(:footnote_entries) ? container.footnote_entries : []
      entry = Array(entries).find { |e| e.id.to_s == id.to_s }
      return "" unless entry

      Array(entry.paragraphs).map { |p| render_inline(p).strip.gsub(/\s+/, " ") }
        .reject(&:empty?).join(" ")
    end

    def drawing_segments(drawing)
      rid = picture_rid(drawing)
      unless rid
        warn_once("drawing without extractable picture (chart/shape) skipped")
        return []
      end
      img = extract_image(rid)
      img ? [Segment.new(raw: "image:#{img[:name]}#{img[:ext]}[]")] : []
    end

    def picture_rid(drawing)
      pic = drawing.inline&.graphic&.graphic_data&.picture ||
            drawing.anchor&.graphic&.graphic_data&.picture
      pic&.blip_fill&.blip&.embed
    end

    def extract_image(rid, name = nil)
      rel = @rels[rid]
      unless rel
        warn_once("relationship #{rid} for image not found")
        return nil
      end
      target = URI.decode_www_form_component(rel[:target])
      zip_path = target.start_with?("/") ? target[1..] : "word/#{target}"
      entry = @zip.find_entry(zip_path)
      unless entry
        warn_once("media entry #{zip_path} not in package")
        return nil
      end
      @fig_seq += 1
      name ||= "image-#{@fig_seq}"
      img = { name: name, ext: File.extname(target),
              data: entry.get_input_stream.read }
      @images << img
      img
    end

    def join_segments(segs, cell: false)
      merged = []
      segs.each do |s|
        unless s.is_a?(Segment)
          str = s.to_s
          merged << Segment.new(raw: str) unless str.empty?
          next
        end
        next if s.raw.nil? && (s.text.nil? || s.text.empty?)

        if s.raw
          merged << s
        elsif (prev = merged.last) && prev.raw.nil? &&
              prev.bold == s.bold && prev.italic == s.italic &&
              prev.sup == s.sup && prev.sub == s.sub
          prev.text = prev.text.to_s + s.text.to_s
        else
          merged << s.dup
        end
      end
      merged.map { |s| render_segment(s, cell: cell) }.join
    end

    def render_segment(s, cell: false)
      return s.raw.to_s if s.raw

      t = s.text.to_s
      return "" if t.strip.empty?

      t = t.tr("|", "❘") if cell && t.include?("|")
      lead = t[/\A[ \s]*/]
      trail = t[/[ \s]*\z/]
      core = t.strip
      core = core.gsub(" ", "{nbsp}") if (s.sup || s.sub) && core.include?(" ")
      core = "^#{core}^" if s.sup
      core = "~#{core}~" if s.sub
      core = "*#{core}*" if s.bold
      core = "_#{core}_" if s.italic
      "#{lead}#{core}#{trail}"
    end

    # -------------------------------------------------- block rendering

    def render_blocks(blocks, out)
      prev_type = nil
      blocks.each do |b|
        consecutive_list = b.type == :list_item && prev_type == :list_item
        out << "" if !out.empty? && !consecutive_list &&
                    !%i[skip source_line].include?(prev_type)
        send("render_#{b.type}", b, out)
        prev_type = b.type
      end
    end

    def render_skip(_b, _out); end

    # Captions that did not attach to a figure/table fall back to body text.
    def render_caption(b, out)
      out << hardbreaks(render_inline(b.para))
    end

    def render_preface(_b, out)
      out << "[.preface]"
      out << "== Document Control"
    end

    def render_heading(b, out)
      title = render_inline(b.para).strip
      if b.biblio
        out << "[bibliography]"
      elsif b.is_appendix
        attrs = %w[appendix]
        attrs << "obligation=#{b.obligation}" if b.obligation
        out << "[#{attrs.join(',')}]"
      end
      out << "[[#{b.anchor}]]"
      out << "#{'=' * b.level} #{title}"
    end

    def render_body(b, out)
      text = render_inline(b.para)
      return if text.strip.empty? && b.anchors.empty?

      b.anchors.each { |id| out << "[[#{id}]]" }
      out << hardbreaks(text) unless text.strip.empty?
    end

    def render_note(b, out)
      text = render_inline(b.para)
      text = text.sub(/\A\s*Note\s*\d*\s*[-–—:.]?\s*/i, "")
      return if text.strip.empty?

      b.anchors.each { |id| out << "[[#{id}]]" }
      out << "[NOTE]"
      out << "===="
      out << hardbreaks(text)
      out << "===="
    end

    def render_list_item(b, out)
      marker = b.fmt == "bullet" ? "*" * (b.ilvl + 1) : "." * (b.ilvl + 1)
      text = hardbreaks(render_inline(b.para))
      out << "#{marker} #{text}".rstrip
    end

    def render_source(b, out)
      out << "[source]"
      out << "----"
      Array(b.paras).each do |p|
        source_text(p).split("\n", -1).each do |line|
          # A line of only dashes/dots would terminate the listing fence.
          line = " #{line}" if line.match?(/\A[-.]{4,}\z/)
          out << line.rstrip
        end
      end
      out << "----"
    end

    def source_text(para)
      para.runs.flat_map { |r| source_run_text(r) }.join
    end

    def source_run_text(run)
      parts = []
      each_named(run, RUN_CHILDREN) do |name, item|
        case name
        when "t" then parts << item.content.to_s
        when "br" then parts << "\n"
        when "tab" then parts << "    "
        end
      end
      parts.join
    end

    def render_bib(b, out)
      raw = render_inline(b.para).strip
      raw = raw.sub(/\s*<\s*>\s*\z/, "").gsub(/\s+/, " ")
      footnotes = raw.scan(/footnote:\w+\[[^\]]*\]/).join(" ")
      raw = raw.gsub(/footnote:\w+\[[^\]]*\]/, "").gsub(/\s+/, " ").strip
      raw = raw.gsub(/link:(\S+)\[([^\]]*)\]/) { Regexp.last_match(2) }
      return if raw.empty?

      label, rest = bib_split(raw)
      rest = "#{rest} #{footnotes}".strip
      id = unique_id(slugify(label))
      out << "* [[[#{id},#{label}]]]#{rest.empty? ? '' : " #{rest}"}"
    end

    # Short citation labels in the s102 house style: "IHO S-100",
    # "ISO 19107:2003". The label is taken from the entry prefix.
    def bib_split(raw)
      m = raw.match(/\AISO(?:\/IEC)?\s+[\d.\-]+(?:\s*:\s*\d{4})?/) ||
          raw.match(/\A(?:IHO\s+)?(S-?\d[\d.]*(?:-\d+)?)\s/) ||
          raw.match(/\A(?:IHO|IMO|UN|ITU|IEC|OGC)\s+[A-Za-z]?\-?[\d.]+/)
      if m
        label = m[0].strip.sub(/\A(S-?\d)/) { "IHO #{Regexp.last_match(1)}" }
        label = label.gsub(/\s*:\s*/, ":")
        rest = raw[m[0].size..].to_s.sub(/\A[\s,–—-]*/, "")
        return [label, rest.empty? ? nil : " #{rest}"]
      end

      words = raw.split(" ")
      return [raw, nil] if words.size <= 6

      ["#{words.first(6).join(' ')}…", " #{words[6..].join(' ')}"]
    end

    def render_figure(b, out)
      b.anchors.each { |id| out << "[[#{id}]]" }
      out << "[[#{b.anchor}]]"
      cap = caption_body(b.caption, "figure")
      out << ".#{cap}" unless cap.empty?
      base = cap.empty? ? nil : "figure-#{slugify(cap)}"
      b.images.each_with_index do |d, i|
        name = base ? (i.zero? ? base : "#{base}-#{i + 1}") : nil
        img = extract_image(picture_rid(d), name) if picture_rid(d)
        out << "image::#{img[:name]}#{img[:ext]}[]" if img
      end
    end

    def render_table(b, out)
      tbl = b.table
      out << "[[#{b.anchor}]]"
      cap = caption_body(b.caption, "table")
      out << ".#{cap}" unless cap.empty?
      cols = table_columns(tbl)
      out << %([cols="#{Array.new(cols, 'a').join(',')}"])
      out << "|==="
      tbl.rows.each { |row| out << render_row(row) }
      out << "|==="
    end

    def table_columns(tbl)
      n = Array(tbl.grid&.columns).size
      n = tbl.rows.map { |r| r.cells.sum { |c| colspan(c) } }.max if n.zero? || n > 30
      [n, 1].max
    end

    def colspan(cell)
      c = cell.properties&.grid_span&.value.to_i
      c.zero? ? 1 : c
    end

    def render_row(row)
      cells = []
      skip = 0
      row.cells.each do |cell|
        if skip.positive?
          skip -= 1
          next
        end
        skip = colspan(cell) - 1
        cells << render_cell(cell)
      end
      cells.join
    end

    def render_cell(cell)
      paras = Array(cell.paragraphs)
      nested = Array(cell.tables)
      header = header_cell?(cell)
      complex = paras.size > 1 || paras.any? { |p| numpr(p) } || nested.any?
      prefix = +""
      prefix << "#{colspan(cell)}+" if colspan(cell) > 1
      prefix << "a" if complex
      prefix << "h" if header && !complex

      body = if nested.any?
               inner = nested.map { |t| nested_table_lines(t) }.join("\n\n")
               parts = paras.map { |p| cell_paragraph(p) }.map(&:strip).reject(&:empty?)
               (parts << inner).join("\n\n")
             elsif complex
               paras.map { |p| cell_paragraph(p) }.map(&:strip).reject(&:empty?).join("\n\n")
             else
               paras.map { |p| render_inline(p, cell: true).strip }
                 .reject(&:empty?).join(" ")
             end
      "#{prefix}|#{body}"
    end

    def cell_paragraph(p)
      num = numpr(p)
      if num
        fmt = list_format(num.first, num.last)
        marker = fmt == "bullet" ? "*" * (num.last + 1) : "." * (num.last + 1)
        "#{marker} #{render_inline(p, cell: true)}"
      else
        hardbreaks(render_inline(p, cell: true))
      end
    end

    def header_cell?(cell)
      paras = Array(cell.paragraphs)
      text = paras.map { |p| p.text }.join.strip
      return false if text.empty?

      paras.all? do |p|
        runs = Array(p.runs)
        next true if runs.empty?

        runs.all? do |r|
          on_off(r.properties&.bold) ||
            Array(r.text).map(&:content).join.strip.empty?
        end
      end
    end

    def nested_table_lines(tbl)
      buf = []
      render_table(Block.new(type: :table, table: tbl), buf)
      buf.join("\n")
    end

    def hardbreaks(text)
      text.split("\n", -1).map do |line|
        # A line of only dashes is an AsciiDoc delimiter; defuse it.
        line.match?(/\A-+\s*\z/) ? " #{line}" : line
      end.join(" +\n")
    end

    # ---------------------------------------------------------- output

    def split_sections
      sections = []
      clause_no = 0
      annex_no = 0
      @blocks.each do |b|
        if b.type == :preface
          sections << { file: "00-document-control.adoc", blocks: [b] }
          next
        end
        if b.type == :heading && b.level == 2
          title = b.para.text.strip
          if b.is_appendix || b.biblio
            annex_no += 1
            letter = ("a".."z").to_a[annex_no - 1] || "z"
            file = "a#{letter}-#{slugify(title)}.adoc"
          else
            clause_no += 1
            file = format("%02d-%s.adoc", clause_no, slugify(title))
          end
          sections << { file: file, blocks: [] }
        elsif sections.empty?
          next if b.type == :skip
          next if b.type == :body && b.para.text.strip.empty? && b.anchors.empty?

          # Content before the first heading belongs to Document Control.
          control = sections.detect { |s| s[:file] == "00-document-control.adoc" }
          next control[:blocks] << b if control

          sections << { file: "00-front-matter.adoc", blocks: [] }
        end
        sections.last[:blocks] << b if sections.last && b.type != :skip
      end
      sections
    end

    def write_output
      sections = split_sections
      FileUtils.mkdir_p(File.join(@out_dir, "sections"))
      FileUtils.mkdir_p(File.join(@out_dir, "images"))
      includes = []
      sections.each do |sec|
        out = []
        render_blocks(sec[:blocks], out)
        content = out.join("\n").gsub(/\n{3,}/, "\n\n").strip + "\n"
        File.write(File.join(@out_dir, "sections", sec[:file]), content)
        includes << sec[:file]
      end
      write_images
      write_document_adoc(includes)
      File.write(File.join(@out_dir, "conversion-report.txt"),
                 @warnings.join("\n") + "\n")
    end

    def write_images
      seen = {}
      @images.each do |img|
        key = "#{img[:name]}#{img[:ext]}"
        next if seen[key]

        seen[key] = true
        File.binwrite(File.join(@out_dir, "images", key), img[:data])
      end
    end

    def header_title
      DOC_CONFIG.fetch(@kind)[:header_title]
    end

    def write_document_adoc(includes)
      header = []
      header << "= #{header_title} (Preview PR)"
      header << ":series: S"
      header << ":docnumber: 121"
      header << ":doctype: standard"
      header << ":edition: 1.1.0"
      header << ":language: en"
      header << ":copyright-year: 2022"
      header << ":committee: hssc"
      header << ":workgroup: s-121pt"
      header << ":toclevels: 3"
      header << ":mn-document-class: iho"
      header << ":mn-output-extensions: xml,html,doc,pdf,rxl"
      header << ":local-cache-only:"
      header << ":data-uri-image:"
      header << ":imagesdir: images"
      annex = DOC_CONFIG.fetch(@kind)[:annex_attr]
      header << ":annex: #{annex}" if annex
      header << ""
      includes.each do |f|
        header << "include::sections/#{f}[]"
        header << ""
      end
      File.write(File.join(@out_dir, "document.adoc"), header.join("\n") + "\n")
    end
  end
end

if $PROGRAM_NAME == __FILE__
  options = { kind: "main" }
  OptionParser.new do |op|
    op.on("--kind KIND", %w[main annex-a annex-b]) { |v| options[:kind] = v }
  end.parse!(ARGV)

  if ARGV.size != 2
    warn "usage: #{File.basename($0)} input.docx OUT_DIR [--kind main|annex-a|annex-b]"
    exit 1
  end

  input, out_dir = ARGV
  FileUtils.mkdir_p(out_dir)
  converter = DocxToMn::Converter.new(input, out_dir, options[:kind])
  converter.run
  warn "#{converter.warnings.size} warning(s); see #{File.join(out_dir, 'conversion-report.txt')}"
end
