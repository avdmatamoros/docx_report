module DocxReport
  module DataItem
    attr_reader :fields, :images

    def add_field(name, value, type = :text)
      field = Field.new name, value, type
      raise 'duplicate field name' if @fields.any? { |f| f.name == field.name }
      @fields << field
    end
  end
end
