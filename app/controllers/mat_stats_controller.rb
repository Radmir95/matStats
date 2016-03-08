class MatStatsController < ApplicationController
  def index
    if !signed_in?
      redirect_to root_url
    end
  end

  def registration
    @user = User.new
	end

	def newUser
    @user = User.new(:first_name => params[:enterUser][:first_name],
                     :last_name => params[:enterUser][:last_name],
                     :email => params[:enterUser][:email],
                     :password => params[:enterUser][:password],
                     :birth_date => params[:enterUser][:birth_date]
    )
    if @user.save
      redirect_to root_url
    else
      render :registration
    end
	end

  def forget

  end

  def processDate
    require 'rubyXL'

    workbook = RubyXL::Workbook.new

    #wtf???
    if !params[:files][:file1].nil?
    sheet1(workbook)
    end
    if !params[:files][:file2].nil?
      sheet2(workbook)
    end
    if !params[:files][:file3].nil?
      sheet3(workbook)
    end
    if (!params[:files][:file4].nil?)
      sheet4(workbook)
    end
    if !params[:files][:file5].nil?
      sheet5(workbook)
    end
    if !params[:files][:file6].nil?
      sheet6(workbook)
    end
    if !params[:files][:file7].nil?
      sheet7(workbook)
    end
    if !params[:files][:file8].nil?
      sheet8(workbook)
    end
    if !params[:files][:file9].nil?
      sheet9(workbook)
    end
    if !params[:files][:file10].nil?
      sheet10(workbook)
    end
    if !params[:files][:file11].nil?
      sheet11(workbook)
    end
    if !params[:files][:file12].nil?
      sheet12(workbook)
    end
    if !params[:files][:file13].nil?
      sheet13(workbook)
    end
    if !params[:files][:file14].nil?
      sheet14(workbook)
    end
    if !params[:files][:file15].nil?
      sheet15(workbook)
    end
    if !params[:files][:file16].nil?
      sheet16(workbook)
    end

    workbook.write("#{current_user.email}.xlsx")
    send_file("#{current_user.email}.xlsx")

  end

  def countFrequency(arrDate, arrInter)
    arrFreq = Array.new
    arrInter.each_index { |index|
      arrFreq.push(0)
    }
    arrFreq.push(0)

    if arrInter.length == 1

      arrDate.each_with_index { |value,index |
        if value.gsub(',','.').to_f < arrInter[0]
          arrFreq[0]+=1
        else
          arrFreq[1]+=1
          end
      }
    else
      arrDate.each_with_index { |value, index|
        if value.gsub(',','.').to_f < arrInter[0]
          arrFreq[0]+=1
          elsif value.gsub(',','.').to_f >= arrInter[arrInter.length-1]
            arrFreq[arrInter.length]+=1
        else
          count = 0
          length = arrInter.length-1
          while count < (length)
            if (value.gsub(',','.').to_f >= arrInter[count]) && (value.gsub(',','.').to_f < arrInter[count+1])
              arrFreq[count+1]+=1
            end
            count+=1
          end
          end
      }

        end
    return arrFreq
    end

  private def sheet1(workbook)
    arr = Array.new
    uploaded_io = params[:files][:file1]
    File.open(uploaded_io.open()) { |f|
      arr.append(f.read.split())
    }

    worksheet = workbook[0]
    worksheet.sheet_name = 'Задание 1'
    worksheet.add_cell(0,0, "Выборка")
    worksheet.change_column_width(1, 23)
    arr[0].each_with_index { |value, index|
      worksheet.add_cell(index+1,0, value.gsub(',','.').to_f)
    }

    worksheet.add_cell(1,1, 'Выборочное среднее = ')
    worksheet.add_cell(2,1, 'Выборочная дисперсия = ')
    worksheet.add_cell(3,1, 'Станд. отклонение = ')
    worksheet.add_cell(4,1, 'Коэфф. ассиметрии = ')
    worksheet.add_cell(5,1, 'Медиана = ')
    worksheet.add_cell(6,1, 'Эксцесс = ')
    worksheet.add_cell(7,1, 'Объем выборки = ')
    worksheet.add_cell(8,1, 'Минимальное значение = ')
    worksheet.add_cell(9,1, 'Максимальное значение = ')
    worksheet.add_cell(10,1, 'Размах = ')

    worksheet.add_cell(1, 2, '', 'AVERAGE(A:A)')
    worksheet.add_cell(2, 2, '', 'VAR(A:A)')

    worksheet.add_cell(3, 2, '', 'SQRT(C3)')
    worksheet.add_cell(4, 2, '', 'SKEW(A:A)')
    worksheet.add_cell(5, 2, '', 'MEDIAN(A:A)')
    worksheet.add_cell(6, 2, '', 'KURT(A:A)')
    worksheet.add_cell(7, 2, '', 'COUNT(A:A)')
    worksheet.add_cell(8, 2, '', 'MIN(A:A)')
    worksheet.add_cell(9, 2, '', 'MAX(A:A)')
    worksheet.add_cell(10, 2, '', 'C10 - C9')
  end

  private def sheet2(workbook)
    arr = Array.new
    uploaded_io = params[:files][:file2]
    File.open(uploaded_io.open()) { |f|
      arr.append(f.read.split())
    }

    #naming columns
    worksheet = workbook.add_worksheet('Задание 2')
    worksheet.add_cell(0,0, "Выборка")
    worksheet.add_cell(0,2, 'Границы интервалов')
    worksheet.add_cell(0,3, 'Середины интервалов')
    worksheet.add_cell(0,4, 'Частота')
    worksheet.add_cell(0,5, 'Норм. распр.')

    arrDate = Array.new

    #date
    arr[0].each_with_index { |value, index|
      if (index >2 )
      worksheet.add_cell(index-2,0, value.gsub(',','.').to_f)
      end
    }

    #intervals
    arrIntervals = Array.new

    count = 1
            (arr[0][2].to_i-1).times {
              arrIntervals.push(arr[0][0].gsub(",",".").to_f+arr[0][1].gsub(",",".").to_f*(count-1))
             worksheet.add_cell(count,2, arrIntervals[count-1] )
             count+=1
            }
    worksheet.add_cell(count,2, ">#{arrIntervals[count-2].to_s}" )


    #middle of intervals
            count = 1
    (arr[0][2].to_i-1).times {
      if count == 1
      worksheet.add_cell(count,3, arr[0][0].gsub(",",".").to_f+arr[0][1].gsub(",",".").to_f/2-arr[0][1].gsub(',','.').to_f)
      else
        worksheet.add_cell(count,3, arr[0][0].gsub(",",".").to_f+arr[0][1].gsub(",",".").to_f/2+ arr[0][1].gsub(",",".").to_f*(count-2))
      end
      count+=1
    }
    worksheet.add_cell(count,3, arr[0][0].gsub(",",".").to_f+arr[0][1].gsub(",",".").to_f/2+ arr[0][1].gsub(",",".").to_f*(count-2))

            count = arr[0].length+1
            i = 0

      #frequency formula
            while i <= count do
              #some magic start

              #end some magic
              i+=1
            end

    #Normal
    count = 1

    worksheet.add_cell(0,7, 'Выборочное среднее')
    worksheet.add_cell(2,7, 'Станд. отклон')
    worksheet.add_cell(1,7,'', 'AVERAGE(A:A)')
    worksheet.add_cell(3,7,'', 'STDEV(A:A)')

    (arr[0][2].to_i).times {
      worksheet.add_cell(count,5,'', "NORMDIST(D#{(count+1).to_s},H2,H4,FALSE)")
    count+=1
    }
    arr[0].delete_at(0)
    arr[0].delete_at(0)
    arr[0].delete_at(0)

    arrFreq = countFrequency(arr[0], arrIntervals)
            count = 0
            length = arrFreq.length
            while count < length
              worksheet.add_cell(count+1,4, "#{arrFreq[count]}")
              count+=1
            end


  end

  private def sheet3(workbook)
    arr = Array.new
    uploaded_io = params[:files][:file3]
    File.open(uploaded_io.open) { |f|
      arr.append(f.read.split)
    }
    worksheet = workbook.add_worksheet('Задание 3')
    worksheet.add_cell(0,0, "Выборка")
    worksheet.add_cell(0,1, "ЭФР")
    worksheet.add_cell(0,2, "Норм. распр.")
    worksheet.add_cell(0,3, "Расхождение")
    worksheet.add_cell(0,4, "Макс. расхожд.")
    worksheet.add_cell(0,5, "Среднее")
    worksheet.add_cell(2,5, "Станд. отклон.")
    worksheet.add_cell(1,5,'', 'AVERAGE(A:A)')
    worksheet.add_cell(3,5,'', 'STDEV(A:A)')

    #date
    arr[0].sort!
    count=0
    #date
    arr[0].each_with_index { |value, index|
      if index == 0
        worksheet.add_cell(index+1,0, value.gsub(',','.').to_f)
      elsif index == (arr[0].length-1)
      worksheet.add_cell(index+1+count,0, value.gsub(',','.').to_f)

      else
        worksheet.add_cell(index+1+count,0, value.gsub(',','.').to_f)
        worksheet.add_cell(index+1+(count+1),0, value.gsub(',','.').to_f)
        count+=1
        end
    }

    #emper function
    i=0
    count=arr[0].length*2

    while i < (count-2)
      worksheet.add_cell(i+1,2,'', "NORMDIST(A#{(i+2).to_s},F2,F4,TRUE)")
      worksheet.add_cell(i+1,3,'', "ABS(B#{i+2}-C#{i+2})")
      if (i == 0)
        worksheet.add_cell(1,1, 0)
      elsif(i%2==1)
        worksheet.add_cell(i+1,1,'', "(ROW(B#{(i+2).to_s})-1)/#{(arr[0].length*2-4).to_s}-1/#{(arr[0].length-2).to_s}")
      else
      worksheet.add_cell(i+1,1,'', "(ROW(B#{(i+1).to_s})-1)/#{(arr[0].length*2-4).to_s}")
      end
      i+=1
    end
    worksheet.add_cell(1,4,'', "MAX(D:D)")
  end

  private def sheet4(workbook)
  arr = Array.new
  uploaded_io = params[:files][:file4]
  File.open(uploaded_io.open()) { |f|
    arr.append(f.read.split())
  }
  alpha = arr[0][0]
  arrInter = Array.new
  arr[0].delete_at(0)
  worksheet = workbook.add_worksheet('Задание 4')
  worksheet.add_cell(1,1, "#{alpha}")
  arr[0].each_with_index { |value, index|
    if (index != 0 && index != 1 && index != 2)
    worksheet.add_cell(index-2,0, value.gsub(',','.').to_f)
    end
  }
  worksheet.add_cell(0,0, "Выборка")
  worksheet.add_cell(0,1, "α")
  worksheet.add_cell(2,1, "H0")
  worksheet.add_cell(3,1, "нормальное")
  worksheet.merge_cells(0, 2, 0, 3)
  worksheet.merge_cells(3, 2, 4, 2)
  worksheet.merge_cells(3, 3, 3, 4)
  worksheet.merge_cells(3, 5, 4, 5)
  worksheet.add_cell(3,2, "границы")
  worksheet.add_cell(3,3, "Частоты")
  worksheet.add_cell(4,3, "выборочные")
  worksheet.add_cell(4,4, "ожидаемые")
  worksheet.add_cell(0, 2, 'Группированые')
  worksheet.add_cell(3, 5, 'χ2 = (vi - npi)^2')
  worksheet.add_cell(1, 2, 'среднее')
  worksheet.add_cell(2, 2, 'станд.отклон.')
  worksheet.add_cell(1, 3, '','AVERAGE(A:A)')
  worksheet.add_cell(2, 3, '','STDEVP(A:A)')
  count = 1
  while count <= arr[0][2].to_f
    if (count == arr[0][2].to_f )
      worksheet.add_cell(count+4,2, ">#{arrInter[count-2].to_s}")
    elsif
      arrInter.push(arr[0][0].gsub(",",".").to_f+arr[0][1].gsub(",",".").to_f*(count-1))
      worksheet.add_cell(count+4,2, arrInter[count-1])
    end
    count+=1
  end
  worksheet.add_cell(count+4,2, 'Σ')
  arr[0].delete_at(0)
  arr[0].delete_at(0)
  arr[0].delete_at(0)
  arrFreq = Array.new
  arrFreq = countFrequency(arr[0], arrInter)
            arrFreq.each_with_index{|value, index|
              worksheet.add_cell(index + 5, 3, value)
            }
  worksheet.add_cell(4, 6, 'Ф(x)')
  worksheet.add_cell(4, 7, 'p')
  arrInter.each_with_index{|value,index|
    worksheet.add_cell(index + 5, 6, '', "NORMDIST(#{value.to_s},D2,D3,TRUE)")
  }
            count = 0
            length = arrInter.length
            while count <= length
              if count==0
                worksheet.add_cell(count + 5, 7,'',"G#{count+6}")
              elsif count == length
                worksheet.add_cell(count + 5, 7,'', "1-G#{count+5}")
              else
                worksheet.add_cell(count + 5, 7,'', "G#{count+6}-G#{count+5}")
              end
              count+=1
            end

            worksheet.add_cell(0, 4,'Объем выборки')
            worksheet.add_cell(1, 4,'',"COUNT(A:A)")
            count = 0
            length = arrInter.length
            while count <= length
              worksheet.add_cell(count + 5, 4,'',"H#{count+6}*E2")
              count+=1
            end

            #summa
              worksheet.add_cell(length+9,3,'',"SUM(D6:D#{(5+length-3)}")
              worksheet.add_cell(length+9,4,'',"SUM(E6:E#{(5+length-3)}")


  end

  private def sheet5(workbook)
    arr = Array.new
    uploaded_io = params[:files][:file3]
    File.open(uploaded_io.open()) { |f|
      arr.append(f.read.split())
    }
    worksheet = workbook.add_worksheet('Задание 5')
    worksheet.add_cell(0,0, "Выборка")
  end

  private def sheet6(workbook)
    arr = Array.new
    uploaded_io = params[:files][:file3]
    File.open(uploaded_io.open()) { |f|
      arr.append(f.read.split())
    }
    worksheet = workbook.add_worksheet('Задание 6')
    worksheet.add_cell(0,0, "Выборка")
  end

  private def sheet7(workbook)
    arr = Array.new
    uploaded_io = params[:files][:file3]
    File.open(uploaded_io.open()) { |f|
      arr.append(f.read.split())
    }
    worksheet = workbook.add_worksheet('Задание 7')
    worksheet.add_cell(0,0, "Выборка")
  end

  private def sheet8(workbook)
    arr = Array.new
    uploaded_io = params[:files][:file3]
    File.open(uploaded_io.open()) { |f|
      arr.append(f.read.split())
    }
    worksheet = workbook.add_worksheet('Задание 8')
    worksheet.add_cell(0,0, "Выборка")
  end

  private def sheet9(workbook)
    arr = Array.new
    uploaded_io = params[:files][:file3]
    File.open(uploaded_io.open()) { |f|
      arr.append(f.read.split())
    }
    worksheet = workbook.add_worksheet('Задание 9')
    worksheet.add_cell(0,0, "Выборка")
  end

  private def sheet10(workbook)
    arr = Array.new
    uploaded_io = params[:files][:file3]
    File.open(uploaded_io.open()) { |f|
      arr.append(f.read.split())
    }
    worksheet = workbook.add_worksheet('Задание 10')
    worksheet.add_cell(0,0, "Выборка")
  end

  private def sheet11(workbook)
    arr = Array.new
    uploaded_io = params[:files][:file3]
    File.open(uploaded_io.open()) { |f|
      arr.append(f.read.split())
    }
    worksheet = workbook.add_worksheet('Задание 11')
    worksheet.add_cell(0,0, "Выборка")
  end

  private def sheet12(workbook)
    arr = Array.new
    uploaded_io = params[:files][:file3]
    File.open(uploaded_io.open()) { |f|
      arr.append(f.read.split())
    }
    worksheet = workbook.add_worksheet('Задание 12')
    worksheet.add_cell(0,0, "Выборка")
  end

  private def sheet13(workbook)
    arr = Array.new
    uploaded_io = params[:files][:file3]
    File.open(uploaded_io.open()) { |f|
      arr.append(f.read.split())
    }
    worksheet = workbook.add_worksheet('Задание 13')
    worksheet.add_cell(0,0, "Выборка")
  end

  private def sheet14(workbook)
    arr = Array.new
    uploaded_io = params[:files][:file3]
    File.open(uploaded_io.open()) { |f|
      arr.append(f.read.split())
    }
    worksheet = workbook.add_worksheet('Задание 14')
    worksheet.add_cell(0,0, "Выборка")
  end

  private def sheet15(workbook)
    arr = Array.new
    uploaded_io = params[:files][:file3]
    File.open(uploaded_io.open()) { |f|
      arr.append(f.read.split())
    }
    worksheet = workbook.add_worksheet('Задание 15')
    worksheet.add_cell(0,0, "Выборка")
  end

  private def sheet16(workbook)
    arr = Array.new
    uploaded_io = params[:files][:file3]
    File.open(uploaded_io.open()) { |f|
      arr.append(f.read.split())
    }
    worksheet = workbook.add_worksheet('Задание 16')
    worksheet.add_cell(0,0, "Выборка")
  end



end
