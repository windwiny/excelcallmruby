
source_files = Rake::FileList.new("*.rb") do |fl|
  fl.exclude("~*")
  fl.exclude(/^scratch\//)
end

task :default => :mrb
task :mrb => source_files.ext(".mrb")

rule ".mrb" => ".rb" do |t|
  cmd = "mrbc.exe #{t.source}"
  puts cmd
  system cmd
end

