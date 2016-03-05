require 'tempfile'

module DocxReport
  class Report
    attr_reader :fields, :tables

    def initialize(template_path)
      @template_path = template_path
      @fields = {}
      @tables = []
    end

    def add_field(name, value)
      @fields["{@#{name}}"] = value
    end

    def add_table(name, collection, has_header = false)
      table = Table.new name, has_header
      @tables << table
      yield table
      table.load_records collection
    end

    def generate_docx(filename = nil, template_path = nil)
      doc = Document.new template_path || @template_path
      parser = Parser.new doc
      parser.fill_all_tables @tables
      parser.replace_all_fields @fields
      docx_path = filename || Tempfile.new('')
      doc.save docx_path
      File.read docx_path if filename.nil?
    end
  end
end
