# frozen_string_literal: true

require 'forwardable'
require 'tempfile'
require 'zlib'

module Embulk
  module Output
    class GoogleSpreadsheets < OutputPlugin

      class WriteBuffer
        extend Forwardable
        include Enumerable

        MAX_LEN = 2 ** 32 - 1

        d_crc32 = 'V1'
        d_marshal_length = 'V1'
        RECORD_META_PACK_DIRECTIVES = [d_crc32, d_marshal_length].join

        RECORD_META_SIZE = 4 + 4

        attr_reader :file

        def initialize
          @file = Tempfile.new(safe_class_name)
          @file.binmode
        end

        def close
          return if file.nil?

          file.close
          file.unlink
          file = nil
        end

        def write_record(r)
          # buffer file - record format
          #
          # | Offset | Length | Expl
          # |  0     | 4      | crc32 of marshal data
          # |  8     | 4      | length of marshal data
          # | 12     | ??     | marshal data (maximum size 2^32 − 1)
          m_data  = Marshal.dump(r)
          m_crc32 = Zlib.crc32(m_data)
          m_len   = m_data.size

          if MAX_LEN < m_len
            raise StandardError.new("record size is too large")
          end

          write [m_crc32, m_len].pack(RECORD_META_PACK_DIRECTIVES)
          write m_data
        end

        def each
          file.rewind

          yield read_record until file.eof?
        end

        def_delegators :@file, :write

      private
        # retrun a class name safe for filesystem
        def safe_class_name()
          self.class.name.downcase.gsub(/[^\d\w]/, '_')
        end

        def read_record
          crc32, length = read_meta
          m_data = file.read(length)

          if crc32 != Zlib.crc32(m_data)
            raise StandardError.new("broken record. buffer_file location: #{file.path}")
          end

          Marshal.restore(m_data)
        end

        def read_meta
          file.read(RECORD_META_SIZE).unpack(RECORD_META_PACK_DIRECTIVES)
        end

      end
    end
  end
end
