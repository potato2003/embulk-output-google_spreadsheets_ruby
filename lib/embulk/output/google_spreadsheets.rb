require 'google_drive'
require 'googleauth'
require 'google/apis/sheets_v4'

module Embulk
  module Output

    class GoogleSpreadsheets < OutputPlugin
      Plugin.register_output("google_spreadsheets", self)

      DEFAULT_SCOPE = [
       'https://www.googleapis.com/auth/drive',
       'https://spreadsheets.google.com/feeds/'
      ]

      # support config by file path or content which supported by org.embulk.spi.unit.LocalFile
      # json_keyfile:
      #   content: |
      class LocalFile
        # return JSON string
        def self.load(v)
          if v.is_a?(String)
            File.read(v)
          elsif v.is_a?(Hash)
            v['content']
          end
        end
      end

      def self.transaction(config, schema, count, &control)
        # configuration code:
        task = {
          # required
          "json_keyfile"     => config.param("json_keyfile",      LocalFile, nil),
          "spreadsheet_url"  => config.param("spreadsheet_url",   :string),
          "worksheet_title"  => config.param("worksheet_title",   :string),

          # optional
          "auth_method"      => config.param("auth_method",       :string,  default: "authorized_user"), # 'auth_method' or 'service_account'
          "mode"             => config.param("mode",              :string,  default: "append"), # `replace` or `append`
          "header_line"      => config.param("header_line",       :bool,    default: false),
          "start_column"     => config.param("start_column",      :integer, default: 1),
          "start_row"        => config.param("start_row",         :integer, default: 1),
          "null_string"      => config.param("null_string",       :string,  default: ""),
          "default_timezone" => config.param("default_timezone",  :string,  default: "+09:00"),
          "default_timezone_format" => config.param("default_timezone_format", :string,  default: "%Y-%m-%d %H:%M:%S.%6N %z"),
        }

        mode = task["mode"].to_sym
        raise "unsupported mode: #{mode.inspect}" unless [:append, :replace].include? mode

        worksheet = build_worksheet_client(task)

        #
        # prepare to write records
        #
        determine_start_index(worksheet, task, schema, mode)

        if mode == :replace
          clean_previous_records(worksheet, schema, task)
        end

        if task['header_line']
          write_header_line(worksheet, schema, task, mode)
        end

        worksheet.save

        task_reports = yield(task)
        next_config_diff = {}
        return next_config_diff
      end

      def init
        @mode        = task["mode"].to_sym
        @row         = task["row_index"]
        @col         = task["col_index"]
        @null_string = task["null_string"]

        @worksheet = self.class.build_worksheet_client(task)
      end

      def close
      end

      def add(page)
        base_col_index = @col

        page.each do |record|
          record_with_meta = schema.names.zip(schema.types, record)

          record_with_meta.each do |(name, type, value)|
            @worksheet[@row, @col] = format(type, value)

            @col += 1
          end

          @col  = base_col_index
          @row += 1
        end

        @worksheet.save
      end

      def finish
      end

      def abort
      end

      def commit
        task_report = {}
        return task_report
      end

      def self.build_worksheet_client(task)
        key = StringIO.new(JSON.parse(task['json_keyfile']).to_json)

        credentials = case task['auth_method']
        when 'authorized_user'
          Google::Auth::UserRefreshCredentials.make_creds(json_key_io: key, scope: DEFAULT_SCOPE)
        when 'service_account'
          Google::Auth::ServiceAccountCredentials.make_creds(json_key_io: key, scope: DEFAULT_SCOPE)
        else
          raise ConfigError.new("Unknown auth method: #{task['auth_method']}")
        end
 
        GoogleDrive::Session.new(credentials)
          .spreadsheet_by_url(task["spreadsheet_url"])
          .worksheet_by_title(task["worksheet_title"])
      end

      def self.determine_start_index(worksheet, task, schema, mode)
        task["row_index"] = r = task["start_row"]
        task["col_index"] = c = task["start_column"]
        previous_record_exists = false

        if mode == :append
          column_range = c...(c + schema.length)
          next_row_index = last_record_index(worksheet, column_range) + 1

          if r < next_row_index
            previous_record_exists = true
            task["row_index"] = next_row_index
          end
        end

        task["previous_record_exists"] = previous_record_exists
      end

      def self.last_record_index(worksheet, column_range)
        # find last records in column range on worksheet.
        worksheet.cells
          .select {|(_, col_index), value|  not value.empty? and column_range.include? col_index }
          .map    {|(row_index, _), _| row_index }
          .max or 0
      end

      def self.write_header_line(worksheet, schema, task, mode)
        worksheet.update_cells(task["start_row"], task["start_column"], [schema.names])

        task["row_index"] += 1 unless task["previous_record_exists"]
      end

      def self.clean_previous_records(worksheet, schema, task)
        row_range    = task["row_index"]..worksheet.num_rows
        column_range = task["col_index"]...(task["col_index"] + schema.length)

        row_range.each do |r|
          column_range.each do |c|
            worksheet[r, c] = ''
          end
        end
      end

      def format(type, v)
        return @null_string if v.nil?

        case type
        when :timestamp
          zone_offset = task['default_timezone']
          format      = task['default_timezone_format']

          v.dup.localtime(zone_offset).strftime(format)
        when :json
          v.to_json
        else
          v
        end
      end
    end
  end
end
