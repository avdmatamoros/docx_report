require 'zip'
require 'nokogiri'

module DocxReport
  class Document
    attr_reader :template_path, :files

    def initialize(template_path)
      @template_path = template_path
      zip = Zip::File.open(template_path)
      @files = load_files zip
      zip.close
    end

    def save(output_path)
      template = Zip::File.open @template_path
      Zip::OutputStream.open(output_path) do |output|
        template.each { |entry| add_files template, output, entry.name }
      end
      template.close
    end

    private

    def add_files(template_file, output, entry_name)
      output.put_next_entry entry_name
      if @files.keys.include? entry_name
        output.write @files[entry_name].to_xml
      else
        output.write template_file.read(entry_name)
      end
    end

    def content_types_xpath
      "//*[@ContentType = '#{CONTENT_TYPES.join("' or @ContentType='")}']"
    end

    def load_files(zip)
      @files = {}
      content_type_element = Nokogiri::XML zip.read(CONTENT_TYPE_NAME)
      content_type_element.xpath(content_types_xpath).each do |e|
        filename = e['PartName'][1..-1]
        @files[filename] = Nokogiri::XML zip.read(filename)
      end
      @files
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
