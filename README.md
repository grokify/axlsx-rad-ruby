AxlsxRad - Rapid XLSX Writer for Customized Spreadsheets
========================================================

Synopsis
--------

AxlsxRad provides a rapid capability to generate styled spreadsheets of JSON
objects to enable the creation of reports with both per-row and per-cell
highlighting.

This is accomplished by leveraging JsonDoc's encapsulation of JSON objects
with easy retrieval of property names and values as arrays that are used by
AxlsxRad to populate XLSX spreadsheets. This is done by simply providing an
array of JSON property keys and then adding JsonDoc objects to the worksheet.

Axlsx additionally supports rapid styling of spreadsheets by accepting a
set of styles and rules for applying those styles to each row. Both a header
row and data rows can be styled with both default background styles for the
entire row and custom styles on a per cell basis. Per-cell styles are assigned
by JsonDoc property key so there is no need to create arrays of styles per
row and to remember which column is located in which array index - it is all
done automatically.

Customization consists of:

* a list of JSON property keys that are used as columns
* a dictionary of styles by style key
* overridable header and row methods tto provide custom styling rules

Finally, Axlsx provides full access to the Axlsx package, workbook and
worksheet objects so fast customization can be supported in addition to 
standard access.

Installing
----------

Download and install axlsx_rad with the following:

    gem install axlsx_rad

#Examples
---------

    require 'axlsx_rad'

    # Create an AxlsxRad::Config object, often per-worksheet
    # Define worksheet columns. These keys should exist in JsonDoc
    # object / subclass properties.
    columns = [ :first_name, :last_name, :email_address, :github_username ]

    # Define styles
    styles     =  {
      :blue    => { :bg_color => '00CCFF' },
      :green   => { :bg_color => '00FF99' },
      :ltgreen => { :bg_color => '99FF99' },
      :yellow  => { :bg_color => 'FFFF99' },
      :orange  => { :bg_color => 'FFCC66' },
      :white   => { :bg_color => 'FFFFFF' },
      :head    => { :bg_color => '00FFFF', :b => true}
    }

    # Subclass AxlsxRad::Config object. This object has accepts the columns,
    # styles and unique key initialization parameters, setting them to the
    # aColumns, dStylesCfg and xxKeyUnique attributes. It doesn't matter how
    # they are set but they should be set before instantiating a
    # AxlsxRad::Worksheet object. Additionally, the #getStylesForHeader() and
    # #getStylesForDocument( document ) methods should be overridden to return
    # a style hash as shown below.
    #
    # The contract contains the #getStylesForHeader and
    # #getStylesForDocument( JsonDoc::Document ) methods.
    #
    # The style names used in the reponse hash should correspond to the style
    # keys used in the styles configuration hash above.

    class MyWorksheetConfig < AxlsxRad::Config
      def initialize(columns=[],styles={},unique_key=nil)
        @aColumns    = columns
        @dStylesCfg  = styles
        @xxKeyUnique = unique_key
      end
      def getStylesForHeader()
        return { :bgstyle => :head }
      end
      def getStylesForDocument( oDocument=nil )
        dStyles = { :bgstyle => :white, :properties => {} }
        if oDocument.getProp(:github_username)
          dStyles[:bgstyle] = :green
          dstyles[:properties][:github_username] = :yellow
        elsif oDocument.getProp(:email_address)
          dStyles[:bgstyle] = :blue
          dstyles[:properties][:email_address] = :yellow
        end
        return dStyles
      end
    end

    # Instantiate a config object
    config    = MyWorksheetConfig.new( columns, styles )

    # Instantiate a workbook
    workbook  = AxlsxRad::Workbook.new

    # Instantiate a worksheet with name and config
    worksheet = workbook.addWorksheet( "Users", config )

    # Add documents where user is a JsonDoc::Document subclass
    # with properties matching those tested against in Axlsx::Config
    # subclass.

    users.each do |user|
      worksheet.addDocument( user )
    end

    # Write XLSX spreadsheet
    workbook.serialize('example_users.xlsx')

    # Write XLSX spreadsheet by accessing the Axlsx::Package object
    workbook.oAxlsx.serialize('example_users.xlsx')

#Documentation
--------------

This gem is 100% documented with YARD, an exceptional documentation library. To see documentation for this, and all the gems installed on your system use:

    $ gem install yard
    $ yard server -g

Notes
-----

1. Removing duplicate documents
 - If the unique_key parameter is set in the AxlsxRad::Config object, that key will be checked to retain or discard added documents which can streamline the adding process. Not setting this value or setting this value to nil will disable checking and add all documents. 

#Change Log
-----------

- **2014-03-23**: 0.0.2
  - Add ability to filter duplicate documents
  - Add AxlsxRad::Workbook::serialize shortcut
- **2014-03-21**: 0.0.1
  - Initial release

#Links
------

Axlsx

https://rubygems.org/gems/axlsx

JsonDoc

https://rubygems.org/gems/jsondoc

#Copyright and License
----------------------

AxlsxRad &copy; 2014 by [John Wang](mailto:johncwang@gmail.com).

AxlsxRad is licensed under the MIT license. Please see the LICENSE document for more information.

Warranty
--------

This software is provided "as is" and without any express or implied warranties, including, without limitation, the implied warranties of merchantibility and fitness for a particular purpose.