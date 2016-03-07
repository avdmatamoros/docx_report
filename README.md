# docx_report [![Build Status](https://travis-ci.org/abudaqqa/docx_report.svg?branch=master)](https://travis-ci.org/abudaqqa/docx_report) [![Code Climate](https://codeclimate.com/github/abudaqqa/docx_report/badges/gpa.svg)](https://codeclimate.com/github/abudaqqa/docx_report) [![Test Coverage](https://codeclimate.com/github/abudaqqa/docx_report/badges/coverage.svg)](https://codeclimate.com/github/abudaqqa/docx_report/coverage) [![Gem Version](https://badge.fury.io/rb/docx_report.svg)](https://badge.fury.io/rb/docx_report)

light weight gem that generates docx files by replacing strings on
previously created .docx file

### Installation

Add this to your Gemfile and run `bundle install`:

```ruby
gem 'docx_report'
```

### Usage

To generate the report you have to create .docx template using MS word and set
fields inside your template to be replaced with data

fields name should be wrapped with braces and start with @ for example
```
{@name}
```

to generate report based on template with the previous name field
```ruby
report = DocxReport.create_docx_report 'public/template.docx'
report.add_field 'name', 'Ahmed Abudaqqa'
```

You can also set tables inside your template but you need first to give it a
name using MS word, the fields should all set on the first or second row,
it depends if you want to use header

```ruby
report.add_table 'table1', @items do |table|
  table.add_field(:title, :name)
  table.add_field(:description, :more)
end
```

To save the output document

```ruby
send_data report.generate_docx,
  type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  disposition: 'attachment',
  filename: 'output.docx'
```

### TODO

- Improve documentation
- Support inserting images
- Allow more tables customization
