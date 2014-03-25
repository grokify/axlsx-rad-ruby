require 'axlsx'

module AxlsxRad
  class Workbook
    attr_accessor :oAxlsx

    def initialize(oAxlsx=nil)
      @oAxlsx      = oAxlsx || Axlsx::Package.new
      @dWorksheets = {}
    end

    def addWorksheet(sWorksheetName=nil,oWsConfig=nil)
      if @dWorksheets.has_key?(sWorksheetName)
        raise RuntimeError, "E_WORKSHEET_EXISTS [#{sWorksheetName}]"
      end
      oAwsWorksheet = AxlsxRad::Worksheet.new(@oAxlsx,sWorksheetName,oWsConfig)
      @dWorksheets[ sWorksheetName ] = oAwsWorksheet
      return oAwsWorksheet
    end

    def addWorksheetDocument(sWorksheetName=nil,oJsondoc=nil)
      if sWorksheetName.nil? || ! sWorksheetName.kind_of?(String)
        raise ArgumentError, 'E_BAD_WORKSHEET_NAME'
      elsif ! @dWorksheets.has_key?( sWorksheetName )
        raise ArgumentError, 'E_WORKSHEET_DOES_NOT_EXIST'
      end
      @dWorksheets[ sWorksheetName ].addDocument( oJsondoc )
    end

    def serialize(sPath=nil)
      @oAxlsx.serialize(sPath)
    end

  end
end

module AxlsxRad
  class Worksheet
    attr_accessor :oAxlsx
    attr_accessor :oAxlsxWorksheet

    def initialize(oAxlsx=nil,sName=nil,oWsConfig=nil)
      @oAxlsx          = oAxlsx
      unless oWsConfig.is_a?(AxlsxRad::Config)
        raise ArgumentError, 'E_NO_AXLSX_RAD_WORKSHEET_CONFIG'
      end
      @oAxlsxWorksheet = @oAxlsx.workbook.add_worksheet(:name => sName)
      @oWsConfig       = oWsConfig
      self.loadWsConfig( oWsConfig )
      @bHasHead        = false
      @dKeysUnique     = {}
    end

    def loadWsConfig(oWsConfig=nil)
      @aColumns       = oWsConfig.aColumns || []
      @dStylesCfg     = oWsConfig.dStylesCfg.is_a?(Hash) ? oWsConfig.dStylesCfg : {}
      @dStylesAxlsx   = getStylesAxlsx( @oAxlsxWorksheet, @dStylesCfg )
      @dStylesColumns = getStylesColumns( @dStylesAxlsx, @aColumns )
    end

    def addDocument(oJsondoc=nil)
      unless @bHasHead
        aStyles = getStylesForHeader
        aValues = oJsondoc.getDescArrayForProperties( @aColumns )
        if aStyles.is_a?(Fixnum) || aStyles.is_a?(Array)
          @oAxlsxWorksheet.add_row aValues, :style => aStyles
        else
          @oAxlsxWorksheet.add_row aValues
        end
        @bHasHead = true
      end
      unless @oWsConfig.xxKeyUnique.nil?
        xxKeyUniqueValue = oJsondoc.getProp( @oWsConfig.xxKeyUnique )
        if @dKeysUnique.has_key?( xxKeyUniqueValue )
          return
        else
          @dKeysUnique[ xxKeyUniqueValue ] = 1
        end
      end
      aValues = oJsondoc.getValArrayForProperties( @aColumns )
      aStyles = getStylesForDocument( oJsondoc )
      if aStyles.is_a?(Fixnum) || aStyles.is_a?(Array)
        @oAxlsxWorksheet.add_row aValues, :style => aStyles
      else
        @oAxlsxWorksheet.add_row aValues
      end
    end

    private

    def symbolizeArray(aAny=[])
      (0..aAny.length-1).each do |i|
        aAny[i] = aAny[i].to_sym if aAny[i].is_a?(String)
      end
      return aAny
    end

    def getStylesAxlsx(oAxlsxWorksheet=nil,dStylesCfg=nil)
      dStylesAxlsx = {}
      dStylesCfg.keys.each do |yStyle|
        dStylesAxlsx[ yStyle ] = oAxlsxWorksheet.styles.add_style( @dStylesCfg[ yStyle ] )
      end
      return dStylesAxlsx
    end

    def getStylesColumns(dStylesAxlsx=nil,aColumns=nil)
      dStylesColumns = {}
      dStylesAxlsx.keys.each do |yStyle|
        dStylesColumns[yStyle] ||= []
        oStyleAxlsx = dStylesAxlsx[yStyle]
        (0..aColumns.length-1).each do |i|
          dStylesColumns[yStyle].push( oStyleAxlsx )
        end
      end
      return dStylesColumns
    end

    def getStylesForHeader()
      aStyles = nil
      if @oWsConfig.is_a?(AxlsxRad::Config)
        dStyles = @oWsConfig.getStylesForHeader
        aStyles = getStylesArrayForHash( dStyles )
      end
      return aStyles
    end

    def getStylesForDocument(oJsondoc=nil)
      aStyle = nil
      if @oWsConfig.is_a?(AxlsxRad::Config)
        dStyle = @oWsConfig.getStylesForDocument(oJsondoc)
        aStyle = getStylesArrayForHash( dStyle )
      end
      return aStyle
    end

    def getStylesArrayForHash(dStyles={})
      if dStyles.is_a?(Hash) && dStyles.keys.length == 0
        return nil
      elsif ! dStyles.is_a?(Hash)
        return nil
      end
      yStyleBg = dStyles.has_key?(:bgstyle) ? dStyles[:bgstyle] : nil
      yStyleBg = yStyleBg.to_sym if yStyleBg.is_a?(String)
      aStyles  = []
      if dStyles.has_key?(:properties) && dStyles[:properties].is_a?(Hash) \
      && dStyles[:properties].keys.length > 0
        @aColumns.each do |yColumn|
          if dStyles[:properties].has_key?(yColumn)
            yStyleCell = dStyles[:properties][yColumn]
            yStyleCell = yStyleCell.to_sym if yStyleCell.is_a?(String)
            aStyles.push( @dStylesAxlsx[ yStyleCell ] )
          elsif yStyleBg
            aStyles.push( @dStylesAxlsx[ yStyleBg ] ) 
          else
            aStyles.push( nil )
          end
        end
      elsif yStyleBg
        aStyles = @dStylesColumns[yStyleBg]
      else
        aStyles = nil
      end
      return aStyles
    end

  end
end