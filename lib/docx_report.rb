require 'docx_report/report'
require 'docx_report/document'
require 'docx_report/parser'
require 'docx_report/table'
require 'docx_report/record'

module DocxReport
  def self.create_docx_report(template_path)
    Report.new template_path
  end
end
