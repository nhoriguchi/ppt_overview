# https://qiita.com/www-tacos/items/4246d5344188d1caa51a
# これ Box の共有リンクを自動的に得る方法を示唆している。

require 'fileutils'
require 'yaml'
require 'optparse'
require 'digest'

outdir = "tmp_out"
FileUtils.mkdir_p(outdir)

# 特定のディレクトリ配下を検索する、というモードと、直接ファイルを指定するモードを用意する。
# 相対パスを表示する (Box Drive の場合、Box 先頭からの相対パス、ただこれ人によって見える範囲が異なる可能性がある。)
# リンクを貼る、特に Box の共有 URL を見ておきたい
# 構造がわかるようにする、これはインプットファイルでグルーピングや構造を制御
# DONE インプットファイルを解釈できるようにする
# DONT tmp_dir 配下の名前衝突問題、ID をふって回避するか
# DONE 面倒だから sort by timestamp にするのがよさそうだ。
# 絶対パス名から tmp_dir配下のハッシュを得る。ファイル名重複を回避するため。

# インプットの設計をどうするか
# YAML かな。

options = {
  :outdir => "tmp_out"
}

# OptionParser の設定
opt = OptionParser.new do |opts|
  opts.banner = "Usage: ruby #{__FILE__} [options]"

  opts.on("-f", "--file config_file", "YAML ファイル") do |file|
    options[:config] = YAML.load(File.read(file))
  end

  opts.on("-v", "--[no-]verbose", "詳細出力を有効にする") do |v|
    options[:verbose] = v
  end

  opts.on("-h", "--help", "ヘルプを表示") do
    puts opts
    exit
  end
end

# パース実行
begin
  opt.parse!
rescue OptionParser::InvalidOption => e
  puts e
  puts opt
  exit 1
end

# オプションの内容を表示（デバッグ用）
puts options.inspect

class PPT
  attr_accessor :mtime

  def initialize path, options
    @options = options
    @digest = Digest::MD5.hexdigest(File.expand_path(path))[0,6]
    dir = File.dirname(File.expand_path(path))
    @bname_nodig = File.basename(path, File.extname(path))
    @bname = File.basename(path, File.extname(path)) + "-#{@digest}"

    @mtime = @dir_mtime = File.mtime("#{path}")
    if Dir.exist? "#{@options[:outdir]}/#{@bname}"
      @dir_mtime = File.mtime("#{@options[:outdir]}/#{@bname}")
    end

    if Dir.exist? "#{@options[:outdir]}/#{@bname}"
      if @mtime > @dir_mtime
        # regenerate image files to update to latest images
        FileUtils.rm_rf("#{@options[:outdir]}/#{@bname}")
      else
        return
      end
    end

    p "generate image files for #{path}"
    cmd = "powershell -File ppt2png.ps1 \"#{path}\""
    system cmd
    p "moved #{dir}/#{@bname_nodig} to #{@options[:outdir]}"
    FileUtils.mv("#{dir}/#{@bname_nodig}", "#{@options[:outdir]}/#{@bname}")
    File.write("#{@options[:outdir]}/#{@bname}/path.txt", "#{File.expand_path(path)}")
    File.write("#{@options[:outdir]}/#{@bname}/timestamp.txt", "#{File.mtime(path)}")
  end

  def print_md
    dir = "#{@options[:outdir]}/#{@bname}"
    one_section_mdtext = []

    # TODO: レベルを外から制御できたりしないか。
    one_section_mdtext.push "## #{@bname_nodig}\n"

    metadata = []
    if File.exist? "#{dir}/path.txt"
      path = File.read("#{dir}/path.txt")
      metadata.push "- File path: #{path}"
    end
    if File.exist? "#{dir}/timestamp.txt"
      tstamp = File.read("#{dir}/timestamp.txt")
      metadata.push "- Timestamp: #{tstamp}"
    end
    if ! metadata.empty?
      one_section_mdtext.push(metadata.join("\n") + "\n")
    end

    1.upto(1000) do |i|
      break unless File.exist? "#{@options[:outdir]}/#{@bname}/#{i}.png"
      one_section_mdtext.push "<img src='#{@options[:outdir]}/#{@bname}/#{i}.png' width='32%' />"
    end

    return one_section_mdtext.join("\n")
  end
end

ppts = []
mdtext = []
options[:config]["groups"].each do |group|
  if group["type"] == "file"
    group_ppts = []
    group["path"].each do |f|
      next unless f =~ /pptx$/
      next if f =~ /\/~/
      next unless File.exist? f

      ppt = PPT.new f, options
      group_ppts.push ppt
    end
    group_ppts.sort! do |a, b|
      b.mtime <=> a.mtime
    end
    mdtext.push("# group: #{group["name"]}")
    group_ppts.each do |ppt|
      mdtext.push ppt.print_md
    end
  elsif group["type"] == "directory"
    # p Dir.glob("#{group["path"]}/**/*.pptx")
    group_ppts = []
    Dir.glob("#{group["path"]}/**/*.pptx").each do |f|
      next unless File.exist? f
      next if f =~ /\/~/

      ppt = PPT.new f, options
      group_ppts.push ppt
    end
    group_ppts.sort! do |a, b|
      b.mtime <=> a.mtime
    end
    mdtext.push("# group: #{group["name"]}")
    group_ppts.each do |ppt|
      mdtext.push ppt.print_md
    end
  end
end
File.write("./out.md", mdtext.join("\n\n---\n"))
