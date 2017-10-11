# Excel2JSON-Demo
Excel2JSON-Demo,parse Excel Data to JSON Data

## 集成libxls
### 1.放到根目录
### 2.制作静态库XXX
### 3.XXX静态库 header search paths设置为
./libxls
./libxls/include
./libxls/include/libxls

### 4.main target的target dependencies选择XXX
### 5.从producet中拖XXX到main target的link binary with libraries中
### 6.main target导入libiconv
