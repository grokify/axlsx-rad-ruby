module AxlsxRad
  class Config
    attr_accessor :aColumns
    attr_accessor :dStylesCfg
    attr_accessor :xxKeyUnique
    def initialize(aColumns=[],dStylesCfg={},xxKeyUnique=nil)
      @aColumns    = aColumns
      @dStylesCfg  = dStylesCfg
      @xxKeyUnique = nil
    end
    # Override this method with rules to return style key
    def getStylesForHeader()
      return { :bgstyle => :head }
    end
    def getStylesForDocument(oJsondoc=nil)
      dStyles = { :bgstyle => nil, :properties => {} }
      unless oJsondoc.is_a?(JsonDoc::Document)
        raise ArgumentError, 'E_NOT_JSONDOC'
      end
      return dStyles
    end
  end
end