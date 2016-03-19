module DocxReport
  class Table
    attr_reader :name, :has_header, :records

    def initialize(name, has_header = false)
      @name = name
      @has_header = has_header
      @records = []
      @fields = []
    end

    def new_record
      new_record = Record.new
      records << new_record
      new_record
    end

    def add_field(name, value = nil, type = :text, &block)
      field = Field.new(name, value || block, type)
      raise 'duplicate field name' if @fields.any? do |f|
        f.name == field.name
      end
      @fields << field
    end

    def load_records(collection)
      collection.each do |item|
        record = new_record
        @fields.each do |field|
          record.fields << field.load_value(item)
        end
      end
    end
  end
end
