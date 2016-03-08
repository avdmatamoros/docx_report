module DocxReport
  class Table
    attr_accessor :name, :has_header, :records

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

    def add_field(name, mapped_field = nil, &block)
      @fields << { name: name, mapped_field: mapped_field, block: block }
    end

    def load_records(collection)
      collection.each do |item|
        record = new_record
        @fields.each do |field|
          if field[:block].nil?
            record.add_field field[:name],
                             mapped_value(item, field[:mapped_field])
          else
            record.add_field field[:name], field[:block].call(item)
          end
        end
      end
    end

    private

    def mapped_value(item, field_name)
      item.is_a?(Hash) ? item[field_name] : item.send(field_name)
    end
  end
end
