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
          "is_write_header"=> config.param("is_write_header",:bool,    default: false),
          "start_cell"     => config.param("start_cell",     :string,  default: "A1"),
          "null_representation" => config.param("null_representation", :string,  default: ""),
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

        if task['is_write_header']
          write_header(worksheet, schema, task, mode)
        end

        worksheet.save

        task_reports = yield(task)
        next_config_diff = {}
        return next_config_diff
      end

      def self.write_header(worksheet, schema, task, mode)
        r, c = worksheet.cell_name_to_row_col(task['start_cell'])
        worksheet.update_cells(r, c, [schema.names])

        task["row_index"] += 1 unless task["previous_record_exists"]
      end

      def self.clean_previous_records(worksheet, schema, task)
        row_range    = task["row_index"]..worksheet.num_rows
        column_range = task["col_index"]...(task["col_index"] + schema.length)

        row_range.each do |row|
          column_range.each do |col|
            worksheet[row, col] = ''
          end
        end
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

      def self.determine_start_index(worksheet, task, schema, mode)
        start_row_index, start_col_index = worksheet.cell_name_to_row_col(task["start_cell"])

        task["row_index"] = start_row_index
        task["col_index"] = start_col_index

        previous_record_exists = false

        if mode == :append
          column_range = start_col_index...(start_col_index + schema.length)
          next_row_index = last_record_index(worksheet, column_range) + 1

          if start_row_index < next_row_index
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

      def format(type, v)
        return @null_representation if v.nil?
        v
      end
    end
  end
end
