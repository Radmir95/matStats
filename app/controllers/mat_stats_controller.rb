class MatStatsController < ApplicationController
  def index
    if (!signed_in?)
      redirect_to root_url
    end
  end

  def registration
    @user = User.new()
	end

	def newUser
    @user = User.new(:first_name => params[:enterUser][:first_name],
                     :last_name => params[:enterUser][:last_name],
                     :email => params[:enterUser][:email],
                     :password => params[:enterUser][:password],
                     :birth_date => params[:enterUser][:birth_date]
    )
    if (@user.save)
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
    sheet1(workbook)
    sheet2(workbook)
    sheet3(workbook)

    workbook.write("#{current_user.email}.xlsx")
    send_file("#{current_user.email}.xlsx")

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
      worksheet.add_cell(index+1,0, value.to_f)
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

    #date
    arr[0].each_with_index { |value, index|
      if (index >2 )
      worksheet.add_cell(index-2,0, value.to_f)
      end
    }

    #intervals
    count = 1
            arr[0][2].to_i.times {
             worksheet.add_cell(count,2, arr[0][0].gsub(",",".").to_f+arr[0][1].gsub(",",".").to_f*(count-1))
             count+=1
            }

    #middle of intervals
            count = 1
    (arr[0][2].to_i-1).times {
      if (count == 1)
      worksheet.add_cell(count,3, arr[0][0].gsub(",",".").to_f+arr[0][1].gsub(",",".").to_f/2)
      else
        worksheet.add_cell(count,3, arr[0][0].gsub(",",".").to_f+arr[0][1].gsub(",",".").to_f/2+ arr[0][1].gsub(",",".").to_f*(count-1))
      end
      count+=1
    }
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

    (arr[0][2].to_i-1).times {
      worksheet.add_cell(count,5,'', "NORMDIST(D#{(count+1).to_s},H2,H4,FALSE)")

    count+=1
    }
  end

  private def sheet3(workbook)
    arr = Array.new
    uploaded_io = params[:files][:file3]
    File.open(uploaded_io.open()) { |f|
      arr.append(f.read.split())
    }
    worksheet = workbook.add_worksheet('Задание 3')
    worksheet.add_cell(0,0, "Выборка")
    worksheet.add_cell(0,1, "ЭФР")

    #date
    arr[0].sort!
    count=0
    #date
    arr[0].each_with_index { |value, index|
      if (index == 0)
        worksheet.add_cell(index+1,0, value.to_f)
      else
        worksheet.add_cell(index+1+count,0, value.to_f)
        worksheet.add_cell(index+1+(count+1),0, value.to_f)
        count+=1
        end
    }

    #emper function
    i=0
    count=arr[0].length*2

    while i < count
      if (i == 0)
        worksheet.add_cell(1,1, 0)
      elsif(i%2==1)
        worksheet.add_cell(i+1,1,'', "(ROW(B#{i.to_s})-1)/180-1/90")
      else
      worksheet.add_cell(i+1,1,'', "(ROW(B#{(i).to_s})-1)/180")
      i+=1
      end
      end

  end


end
