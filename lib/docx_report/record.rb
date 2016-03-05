module DocxReport
  class Record
    attr_accessor :fields

    def initialize
      @fields = {}
    end

    def add_field(name, value)
      fields["{@#{name}}"] = value
    end
  end
end
