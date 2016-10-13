# dailyreport
日报数据查询发送脚本
            title_sql = {"new_register":"SELECT count(*) from sys_user u WHERE u.create_time >= '%s' and u.create_time < '%s';" % (yesterday,today),
                        }
            key是excel的sheet名字
            value是执行的查询语句，返回的结果会写入当前sheet的单元格
后期会改进定时发送机制，引入scheduler改进定时发送的机制。
因为生成的是一个excel作为日报直接发送不太美观也没有格式，我的临时做法是将生成的excel作为源文件保存，将另外的编辑好美观的日报模板的数据源指向生成的excel文件。
