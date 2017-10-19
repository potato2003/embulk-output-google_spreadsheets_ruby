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
          "start_cell"     => config.param("start_cell",     :string,  default: "A1"),
          "null_representation" => config.param("null_representation", :string,  default: ""),
        }

        task_reports = yield(task)
        next_config_diff = {}
        return next_config_diff
      end

      def init
        json_keyfile = task["json_keyfile"]
        id           = task["spreadsheet_id"]
        gid          = task["worksheet_gid"]
        @start_cell  = task["start_cell"]
        @mode        = task["mode"].to_sym
        @null_representation = task["null_representation"]

        raise "unsupported mode: #{mode.inspect}" unless [:append, :relace].include? @mode

        session = GoogleDrive::Session.from_config(json_keyfile)

        @worksheet = session.spreadsheet_by_key(id).worksheet_by_gid(gid)
      end

      def close
      end

      def add(page)
        row_index, col_index = determine_start_index
        base_col_index = col_index

        page.each do |record|
          record_with_meta = schema.names.zip(schema.types, record)

          record_with_meta.each do |(name, type, value)|
            Embulk.logger.debug("write column to [#{row_index}, #{col_index}]: #{value}")
            @worksheet[row_index, col_index] = format_cell(type, value)

            col_index += 1
          end

          col_index  = base_col_index
          row_index += 1
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

      def determine_start_index
        start_row_index, start_col_index = @worksheet.cell_name_to_row_col(@start_cell)

        if @mode == :append
          column_range = Range.new(start_col_index, start_col_index + schema.length)
          start_row_index = [last_record_index(column_range) + 1, start_row_index].max
        end

        [start_row_index, start_col_index]
      end

      def last_record_index(column_range)
        # find last records in column range on worksheet.
        @worksheet.cells
          .select {|(_, col_index), value|  not value.empty? and column_range.include? col_index }
          .map    {|(row_index, _), _| row_index }
          .max or 0
      end

      def format_cell(type, v)
        return @null_representation if v.nil?
        v
      end
    end
  end
end
