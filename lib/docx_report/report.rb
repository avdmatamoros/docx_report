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

    def add_table(name, collection = nil, has_header = false)
      table = Table.new name, has_header
      @tables << table
      yield table
      table.load_records collection if collection
    end

    def generate_docx(filename = nil, template_path = nil)
      document = Document.new template_path || @template_path
      apply_changes document
      temp = Tempfile.new('output') if filename.nil?
      docx_path = filename || temp.path
      begin
        document.save docx_path
        File.read docx_path if filename.nil?
      ensure
        temp.close! if temp
      end
    end

    private

    def apply_changes(document)
      parser = Parser.new document
      parser.fill_all_tables @tables
      parser.replace_all_fields @fields
    end
  end
end
