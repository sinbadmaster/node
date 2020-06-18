# json2Excel

```npm install json2excel-node --save```

将对象或数组对象转为excel, 暂时只支持了数组嵌套对象，不支持多重嵌套

1. 安装后的使用方式：
```javascript
const j2eUtils = require('json2excel-node');
j2eUtils.json2Excel(savePath, fileName, jsonData, options);
```

2. 参数说明：

|参数名|参数类型|是否必填|描述
|:--|:--|:--|:--
|savePath|string|Y|转换为excel后的保存路径
|fileName|string|Y|转换为excel后的文件名
|jsonData|Object,Array|Y|待转换的数据
|options|Object|N|选项参数，详情见下方表格

options参数说明：
|参数名|参数类型|是否必填|描述
|:--|:--|:--|:--
|headers|Object|N|表头对象，key值为列对应数据，value为表头内容


如传入如下headers，jsonData中的**name**字段将填充到**姓名**列。
```javascript
headers: { name: '姓名' }
```
