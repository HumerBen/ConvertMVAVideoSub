<# 
 Version:3.0
 功能：将微软虚拟学院的字幕处理成字幕.srt文件
 使用：ScriptMain.ps1 SubFile.FullName
 注意：字幕源txt文件需要是utf8编码格式的，utf8无BOM编码格式的话转换后会乱码

 也可以用下面的命令批量执行字幕文件转换
foreach ( $n in (Get-ChildItem "E:\Documents\Windows\C#\*.txt").FullName )
 {
 .\ScriptMain.ps1 $n
 }

#>


# 初始化计数变量
$lineNumber=1

# 从命令行输入文件名
$subFile = $args

# 调试,输入文件名
#$subFile = ".\1.CFundamentalsForAbsoluteBeginnersM01_high.txt"

# 初始化变量
$SubText = $null

# 当文件名变量不为空时执行
if ( $subFile -ne $null )
{
	$orgSubFile = $subFile

	if ( Test-Path $orgSubFile )
	{
		# 如果文件名中带路径则将其过滤出文件名
		$shortSubFile = $orgSubFile.Substring($orgSubFile.LastIndexOf("\")+1)

		# 命名字幕文件名后缀为.srt
		$outSubFile=$shortSubFile.Replace($shortSubFile.Substring($shortSubFile.LastIndexOf(".")+1),"srt")

		# 输出文件路径
		$OutSubPath = $orgSubFile.Substring(0,$orgSubFile.LastIndexOf("\")+1)

		# 如果字幕文件存在则删除
		if ( Test-Path $OutSubPath$outSubFile )
		{
			Remove-Item $OutSubPath$outSubFile
		}

		# 调试，显示文件名
		#Write-Host "文件名为："$outSubFile

		# 读取原文件内容
		foreach ( $lineContent in ( Get-Content $orgSubFile ) )
		{

			# 定义开始时间
			[datetime] $BeginTimeLong = 0

			# 定义结束时间
			[datetime] $EndTimeLong = 0

			# 将每行的内容转成字符串类型
			$lineContentStr = $lineContent.ToString()
			
			# 第一个引号位置
			$QtIndex1 = $lineContentStr.IndexOf('"')

			# 第二个引号位置
			$QtIndex2 = $lineContentStr.IndexOf('"',$QtIndex1+1)

			# 第三个引号位置
			$QtIndex3 = $lineContentStr.IndexOf('"',$QtIndex2+1)

			# 第四个引号位置
			$QtIndex4 = $lineContentStr.IndexOf('"',$QtIndex3+1)

			# 第五个引号位置
			$QtIndex5 = $lineContentStr.IndexOf('"',$QtIndex4+1)

			# 第六个引号位置
			$QtIndex6 = $lineContentStr.IndexOf('"',$QtIndex5+1)

			if ( ($QtIndex1 -gt 0) -and ($QtIndex2 -gt 0) )
			{
				# 截取每行的开始时间ls
				[double] $BeginTimeShort = $lineContentStr.Substring($QtIndex1+1, $QtIndex2-$QtIndex1-2)
				# Write-Host $BeginTimeShort
			}
			
			if ( ($QtIndex5 -gt 0) -and ($QtIndex6 -gt 0))
			{
				# 截取每行的结束时间
				[double] $EndTimeShort = $lineContentStr.Substring($QtIndex5+1, $QtIndex6-$QtIndex5-2)
				# Write-Host $EndTimeShort
			}

			# 转换开始时间
			$BeginTimeLong = $BeginTimeLong.AddSeconds($BeginTimeShort)

			# 开始时间字符串
			$BeginTimeLongStr = ($BeginTimeLong.Hour).ToString() + ":" + ($BeginTimeLong.Minute).ToString() + ":" + ($BeginTimeLong.Second) + ":" + ($BeginTimeLong.Millisecond).ToString()

			# 调试，输出开始时间
			#Write-Host $BeginTimeLongStr

			# 转换结束时间
			$EndTimeLong = $EndTimeLong.AddSeconds($EndTimeShort)

			# 结束时间字符串
			# 最后的时间用“:”分隔
			#$EndTimeLongStr = ($EndTimeLong.Hour).ToString() + ":" + ($EndTimeLong.Minute).ToString() + ":" + ($EndTimeLong.Second) + ":" + ($EndTimeLong.Millisecond).ToString()

			# 最后的时间用“.”分隔
			$EndTimeLongStr = ($EndTimeLong.Hour).ToString() + ":" + ($EndTimeLong.Minute).ToString() + ":" + ($EndTimeLong.Second) + "." + ($EndTimeLong.Millisecond).ToString()

			# 调试，输出结束时间
			#Write-Host $EndTimeLongStr

			# 定义时间轴变量
			$OutTimeLineStr = $BeginTimeLongStr + " --> " + $EndTimeLongStr

			# 字幕文字左边界
			$LeftText = $lineContentStr.IndexOf(">")

			# 字幕文字右边界
			$RightText = $lineContentStr.LastIndexOf("<")

			if ( ($LeftText -gt 0) -and ($RightText -gt 0) )
			{
				# 截取字幕文字
				$SubText = $lineContentStr.Substring($LeftText+1,$RightText-$LeftText-1)
			}
			

			# 输出文件内容
			
			# 显示每行的行号
			Out-File -Encoding utf8 -InputObject $lineNumber -Append -FilePath "$OutSubPath$outSubFile"
			#Write-Host $lineNumber

			# 显示时间轴
			Out-File -Encoding utf8 -InputObject $OutTimeLineStr -Append -FilePath "$OutSubPath$outSubFile"
			#Write-Host $OutTimeLineStr

			# 输出字内容
			Out-File -Encoding utf8 -InputObject $SubText`n -Append -FilePath "$OutSubPath$outSubFile"
			#Write-Host $SubText
			
			# 累加行号
			$lineNumber +=1

		}

		# 输出完成提示
		Write-Host $subFile "文件转换完成" -ForegroundColor Green
	}
	else
	{
		Write-Host "文件不存在"
	}
	
	

	# 处理文件内容
}
else
{
	Write-Host "请输入文件名"
}
