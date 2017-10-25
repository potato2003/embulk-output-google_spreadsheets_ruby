# use the monckey patching, since bundler could not resolve deps.
# https://github.com/gimite/google-drive-ruby/pull/273
#
# ```
# Bundler could not find compatible versions for gem "google-api-client":
#   In Gemfile:
#     embulk-output-google_spreadsheets (>= 0) java depends on
#       google_drive (>= 2.1.6) java depends on
#         google-api-client (< 0.14.0, >= 0.11.0) java
# 
#     embulk-input-bigquery (>= 0) java depends on
#       google-cloud-bigquery (~> 0.29) java depends on
#         google-api-client (~> 0.14.0) java
# ```
module GoogleDrive
  class ApiClientFetcher
    def initialize(authorization)
      @drive = Google::Apis::DriveV3::DriveService.new
      @drive.authorization = authorization
      # Make the timeout virtually infinite because some of the operations (e.g., uploading a large file)
      # can take very long.

      # This is the Max value of Int (32-bit) represented as milliseconds serving as the maximum allowed
      # timeout value that Google's API seems willing accept (possibly due to the Java Socket lib used
      # by httpclient for JRuby).
      t = (2**31 - 1) / 1000
      @drive.client_options.open_timeout_sec = t
      @drive.client_options.read_timeout_sec = t
      @drive.client_options.send_timeout_sec = t
    end
  end
end
 
