module DocxTemplater
  class Block
    attr_reader :context

    def initialize(context)
      @context = context
    end
  end
end
