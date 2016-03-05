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
      zip = Zip::File.open @template_path
      Zip::OutputStream.open(output_path) do |output|
        zip.each do |entry|
          output.put_next_entry entry.name
          if @files.keys.include? entry.name
            output.write @files[entry.name].to_xml
          else
            output.write zip.read(entry.name)
          end
        end
      end
      zip.close()
    end

    private

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
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'
    ].freeze
  end
end
