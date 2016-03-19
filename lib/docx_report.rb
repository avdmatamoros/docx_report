require 'docx_report/report'
require 'docx_report/document'
require 'docx_report/parser'
require 'docx_report/table'
require 'docx_report/record'
require 'docx_report/field'
require 'docx_report/content_file'
require 'docx_report/hyperlink'

module DocxReport
  def self.create_docx_report(template_path)
    Report.new template_path
  end
end
