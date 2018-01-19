# DysmsSender
By wolcen@msn.com

使用阿里云短信服务,按excel清单发短信

配置内容说明
（前面7个为必须项目）
accessKey    -- 阿里云颁发给用户的访问服务所用的密钥ID
accessSecret -- 阿里云颁发给用户的访问服务所用的密钥
signCode     -- 管理控制台中配置的短信签名（状态必须是验证通过）
templateCode -- 管理控制台中配置的审核通过的短信模板的模板CODE（状态必须是验证通过）
rowStart     -- Excel中记录的开始行,数字1-23767
rowEnd       -- Excel中记录的结束行,数字1-23767
recNum       -- Excel中手机号所在列，字母A-Z
模板中参数    -- key为模板中的参数，值为Excel中对应值所在列，字母A-Z
                若不是列，使用直接值，请在前面加@
示例：

```json
{
	"accessKey":"1234567890",
	"accessSecret":"0987654321"
	"rowStart":"1",
	"rowEnd":"12",
	"templateCode":"SMS_12132",
	"signCode":"杭州西湖",
	"recNum":"D",
	"customer":"B",
	"time1":"@2016-11-16 18:30",
	"time2":"@2016-11-16 19:00",
}
```

