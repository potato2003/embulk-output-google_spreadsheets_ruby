require 'google_drive'

module Embulk
  module Output

    class GoogleSpreadsheets < OutputPlugin
      Plugin.register_output("google_spreadsheets", self)

      def self.transaction(config, schema, count, &control)
        # configuration code:
        task = {
          # required
          "json_keyfile"   => config.param("json_keyfile",   :string),
          "spreadsheet_id" => config.param("spreadsheet_id", :string),

          # optional
          "worksheet_gid"  => config.param("worksheet_gid",  :integer, default: 0),
          "mode"           => config.param("mode",           :string,  default: "append"), # available mode are `replace` and `append`
          "is_write_header"=> config.param("is_write_header",:boolean, default: false),
          "null_representation" => config.param("null_representation", :string,  default: ""),
        }

        start_cell = config.param("start_cell", :string,  default: "A1")
        mode       = task["mode"].to_sym

        raise "unsupported mode: #{mode.inspect}" unless [:append, :replace].include? mode

        worksheet = build_worksheet_client(task)
        task["row_index"], task["col_index"] = determine_start_index(worksheet, mode, start_cell, schema)

        task_reports = yield(task)
        next_config_diff = {}
        return next_config_diff
      end

      def init
        @mode        = task["mode"].to_sym
        @row         = task["row_index"]
        @col         = task["col_index"]
        @null_representation = task["null_representation"]

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
        GoogleDrive::Session.from_config(task["json_keyfile"])
          .spreadsheet_by_key(task["spreadsheet_id"])
          .worksheet_by_gid(task["worksheet_gid"])
      end

      def self.determine_start_index(worksheet, mode, start_cell, schema)
        start_row_index, start_col_index = worksheet.cell_name_to_row_col(start_cell)

        if mode == :append
          column_range = Range.new(start_col_index, start_col_index + schema.length)
          start_row_index = [last_record_index(worksheet, column_range) + 1, start_row_index].max
        end

        [start_row_index, start_col_index]
      end

      def self.last_record_index(worksheet, column_range)
        # find last records in column range on worksheet.
        worksheet.cells
          .select {|(_, col_index), value|  not value.empty? and column_range.include? col_index }
          .map    {|(row_index, _), _| row_index }
          .max or 0
      end

      def format(type, v)
        return @null_representation if v.nil?
        v
      end
    end
  end
end
