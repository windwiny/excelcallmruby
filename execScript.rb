#!/usr/bin/env ruby -w

## Can used global variables:     $log   $excel  $wb   $ws


Dir.chdir File.dirname(__FILE__) rescue nil


if RUBY_ENGINE=='mruby'
class Logger
  attr_accessor :level
  def initialize fn=nil, level=1
    @f = fn ? File.open(fn, 'a') : STDERR
    @level = level
  end
  def close
    @f.close
  end

  FATAL=4
  ERROR=3
  WARN=2
  INFO=1
  DEBUG=0
  %w[fatal error warn info debug].each do |x|
    xx = <<-EOS
    def #{x} *msg
      if @level <= #{x.upcase}
        @f.print "\#{Time.now}  \#{msg.join ' '}\n"
        # @f.flush
      end
    end
    EOS
    eval xx
  end
end
else
  require 'logger'
  require 'win32ole'
end



$log = Logger.new 'execScript.log'
$log.level = Logger::DEBUG

$active_wbn = ENV['execScript_active_workbook_fullname']
$active_wsn = ENV['execScript_active_sheet_name']
$log.info "--> PATH: #{Dir.pwd},  ENV wbn/wsn: #{$active_wbn}\##{$active_wsn},  ARGV: #{ARGV}"

def main
  while x = ARGV.shift
    fns = [x, x+'.mrb', x+'.rb']
    fns.delete(x+'.mrb') unless RUBY_ENGINE=='mruby'
    fn = fns.find { |i| File.file? i }
    if File.file? fn
      $log.info "running file: #{fn}"

      begin
        if !$excel
          $excel = excel = WIN32OLE.connect("Excel.Application")
          unless $excel
            $log.error(" ERROR!! Not find $excel")
            return -2
          end
          if $active_wbn && $active_wsn
            $wb = wb = (1..excel.WorkBooks.count).map { |i| excel.WorkBooks[i] }.find { |n| n.FullName == $active_wbn }
            $ws = ws = (1..wb.WorkSheets.count).map { |i| wb.WorkSheets[i] }.find { |n| n.Name == $active_wsn }
            if !$ws
              $log.error(" ERROR!! Not find WorkBook.FullName==#{$active_wbn} and WorkSheet.Name==#{$active_wsn}")
              return -1
            end
          else
            $wb = wb = excel.ActiveWorkbook
            $ws = ws = wb.ActiveSheet
          end
        end
        $log.info "set \$wb,\$ws = #{$wb.FullName},#{$ws.Name}"

        # do_action
        load fn
      rescue => e
        $log.error e
        $log.error e.backtrace.join("\n")
      end
    end
  end
  0
end

rc = main

exit rc.to_i
