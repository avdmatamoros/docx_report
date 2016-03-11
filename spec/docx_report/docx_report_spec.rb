require 'spec_helper'

describe DocxReport do
  it 'creates new docx report' do
    expect(DocxReport.create_docx_report('spec/files/template.docx'))
      .to be_a DocxReport::Report
  end
end
