<p align="center">
  基于EasyExcel所封装的Excel导入导出工具包
</p>



## 说明 | Instructions

- @ExcelColumns 文件导入时，解析Excel数据对应的表头文件下标
- @ExcelProperty EasyExcel自带的注解
- @ExcelIgnore EasyExcel自带的注解



## 演示 | demonstration

```
    @Data
    public class User {
        private Long id;
        
        @ExcelColumns(index = 0)
        private String name;
        
        @ExcelColumns(index = 1)
        private int sex;
        
        @ExcelColumns(index = 2)
        private int age;
        
        private String code;
    }

```


```
    @Service
    public class TestServiceImpl implements TestService {
        
        @Autowired
        private HttpServletResponse response;
        
        @Autowired
        private UserMapper mapper;
        
        /**
          * 文件导入解析
          */
        public void test1(MultipartFile file) {
            List<User> list = ExcelService.importData(file, User.class)
            // TODO
        }
        
        /**
          * 文件导出下载
          */
        public void test2(String fileName) {
            ExcelService.exportData(response, mapper.getList(), fileName)
        }
        
    }  
```


```xml
<dependency>
    <groupId>com.github.Ling2099</groupId>
    <artifactId>simple-excel</artifactId>
    <version>Latest Version</version>
</dependency>
```



## 期望 | Futures

> 欢迎提出更好的意见，帮助完善各个功能

