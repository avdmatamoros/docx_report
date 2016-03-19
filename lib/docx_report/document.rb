require 'zip'
require 'nokogiri'

module DocxReport
  class Document
    attr_accessor :template_path, :files

    def initialize(template_path)
      @template_path = template_path
      zip = Zip::File.open(template_path)
      load_files zip
      zip.close
    end

    def save_to_memory
      Zip::OutputStream.write_buffer do |output|
        add_files output
      end.string
    end

    def save_to_file(output_path)
      Zip::OutputStream.open(output_path) do |output|
        add_files output
      end
    end

    def new_uniqe_id(type)
      (@files.map do |file|
        file.xml.xpath('//*[@id]').map { |e| e[:id].to_i if e.name == type }
      end.flatten.compact.max || 0) + 1
    end

    private

    def add_files(output)
      Zip::File.open @template_path do |template|
        template.each do |entry|
          write_files entry.name, template, output
        end
        @files.each do |file|
          if file.new_rels && file.rels_has_items?
            output.put_next_entry file.rels_name
            output.write file.rels_xml.to_xml
          end
        end
      end
    end

    def write_files(name, template, output)
      if @files.any? { |f| f.name == name }
        add_data name, @files.detect { |f| f.name == name }.xml.to_xml, output
      elsif @files.any? { |f| f.rels_name == name }
        file = @files.detect { |f| f.rels_name == name }
        add_data name, file.rels_xml.to_xml, output if file.rels_has_items?
      else
        add_data name, template.read(name), output
      end
    end

    def add_data(name, data, output)
      output.put_next_entry name
      output.write data
    end

    def content_types_xpath
      "//*[@ContentType = '#{CONTENT_TYPES.join("' or @ContentType='")}']"
    end

    def load_files(zip)
      @files = []
      content_type_node = Nokogiri::XML zip.read(CONTENT_TYPE_NAME)
      content_type_node.xpath(content_types_xpath).each do |e|
        filename = e['PartName'][1..-1]
        @files << ContentFile.new(filename, zip)
      end
    end

    CONTENT_TYPE_NAME = '[Content_Types].xml'.freeze

    CONTENT_TYPES = [
      'application/vnd.openxmlformats-officedocument.wordprocessingml'\
      '.document.main+xml',
      'application/vnd.openxmlformats-officedocument.wordprocessingml'\
      '.header+xml',
      'application/vnd.openxmlformats-officedocument.wordprocessingml'\
      '.footer+xml'
    ].freeze
  end
end
