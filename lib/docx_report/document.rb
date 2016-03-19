require 'zip'
require 'nokogiri'

module DocxReport
  class Document
    attr_accessor :template_path, :files, :image_ids, :images, :image_types

    def initialize(template_path)
      @images = []
      @image_types = []
      @template_path = template_path
      zip = Zip::File.open(template_path)
      load_files zip
      @image_ids = images_ids zip
      zip.close
    end

    def save_to_memory
      Zip::OutputStream.write_buffer do |output|
        add_images output
        add_files output
      end.string
    end

    def save_to_file(output_path)
      Zip::OutputStream.open(output_path) do |output|
        add_images output
        add_files output
      end
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
      if file = @files.detect { |f| f.name == name }
        add_data name, file.xml.to_xml, output
      elsif file = @files.detect { |f| f.rels_name == name }
        add_data name, file.rels_xml.to_xml, output if file.rels_has_items?
      elsif name == CONTENT_TYPE_NAME
        add_data name, @content_types.to_xml, output
      else
        add_data name, template.read(name), output
      end
    end

    def add_data(name, data, output)
      output.put_next_entry name
      output.write data
    end

    def add_images(output)
      images.each do |image|
        image.save(output)
        image.new_rels.each do |rels|
          rels[:Target] = format(rels[:Target], image.type)
        end
        add_content_type image.type
      end
    end

    def content_types_xpath
      "//*[@ContentType = '#{CONTENT_TYPES.join("' or @ContentType='")}']"
    end

    def load_files(zip)
      @files = []
      @content_types = Nokogiri::XML zip.read(CONTENT_TYPE_NAME)
      @content_types.xpath(content_types_xpath).each do |e|
        filename = e['PartName'][1..-1]
        @files << ContentFile.new(filename, zip)
      end
      find_image_types.each { |type| @image_types << type[:Extension] }
    end

    def images_ids(zip)
      zip.entries.map do |e|
        if e.name.start_with?('word/media/image')
          (File.basename(e.name, '.*')[5..-1]).to_i
        end
      end.compact
    end

    def find_image_types
      @content_types.xpath('//*[starts-with(@ContentType, "image")]')
    end

    def add_content_type(type)
      unless @image_types.include? type
        @content_types.children.first << Nokogiri::XML(
          format('<Default Extension="%s" ContentType="image/%s"/>',
                 type, type)).children.first
        @image_types << type
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
