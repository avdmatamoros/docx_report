# docx_report [![Build Status](https://travis-ci.org/abudaqqa/docx_report.svg?branch=master)](https://travis-ci.org/abudaqqa/docx_report) [![Code Climate](https://codeclimate.com/github/abudaqqa/docx_report/badges/gpa.svg)](https://codeclimate.com/github/abudaqqa/docx_report) [![Test Coverage](https://codeclimate.com/github/abudaqqa/docx_report/badges/coverage.svg)](https://codeclimate.com/github/abudaqqa/docx_report/coverage) [![Gem Version](https://badge.fury.io/rb/docx_report.svg)](https://badge.fury.io/rb/docx_report)

light weight gem that generates docx files by replacing strings on
previously created .docx file

### Installation

Add this to your Gemfile and run `bundle install`:

```ruby
gem 'docx_report'
```
### Changes on version 0.2.1
- allow adding image fields
- allow adding nested tables

to insert images you have to make sure ImageMagick is installed, for more
details see https://github.com/minimagick/minimagick

### Important changes on 0.1.0 version
The params on the template document is changed from {@param} to @param@ to
avoid html encoding when using params in urls

### Usage

To generate the report you have to create .docx template using MS word and set
fields inside your template to be replaced with data

fields name should be start and ends with @ for example name
field should look like @name@

images and hyperlink fields need be name as text fields for example @image@ or
@link@

tables names are not required to start or ends with, so 'subitems' is valid name

to title for images inside Microsoft word select the image, right click >>
Picture Format >> Layout & Properties, then you can change the title field

to add table title select the table right click >> Table Properties >>
Alt Text, then you can change the title fields

for hyperlinks you need to set the url as your field name, for example @link@

to generate report based on template with the previous name field
```ruby
report = DocxReport.create_docx_report 'public/template.docx'
report.add_field 'name', 'Ahmed Abudaqqa'
report.add_field 'url', 'http://www.abudaqqa.com', :hyperlink
report.add_image 'picture', 'http://www.abudaqqa.com/avatar.png', :image
```

You can also set tables inside your template and then give it a title using
MS word. and then you can fill it with data by passing a collection of data

```ruby
report.add_table 'table1', @users do |table|
  table.add_field(:title, :name)
  table.add_field(:description) { |user| "details: #{user.details}" }
  table.add_field(:pic_link, :avatar_link, :hyperlink)
  table.add_field(:url, nil, :hyperlink) do |user|
    "http://abudqqa.com/users/id=#{user.id}"
  end

  # nested tables
  subtable1 = table.add_table(:subtable1, :friends)
  subtable.add_field(:name, :name)

  subtable2 = table.add_table(:subtable2) do |item|
    item.contacts.select { |contact| contact.phone }
  end
  subtable2.add_field(:number, :phone)
end
```
In the previous example the first row of the table will be repeated and filled
with the collection data. if you want to leave the first row for header you can
pass true for has_header parameter

```ruby
report.add_table 'table1', @users, true do |table|
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
- Allow more tables customization
- Add more options when insert images
- Support filling charts data
