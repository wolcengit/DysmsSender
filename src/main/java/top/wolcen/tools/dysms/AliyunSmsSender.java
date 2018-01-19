package top.wolcen.tools.dysms;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import com.aliyuncs.DefaultAcsClient;
import com.aliyuncs.IAcsClient;
import com.aliyuncs.dysmsapi.model.v20170525.SendSmsRequest;
import com.aliyuncs.dysmsapi.model.v20170525.SendSmsResponse;
import com.aliyuncs.exceptions.ClientException;
import com.aliyuncs.profile.DefaultProfile;
import com.aliyuncs.profile.IClientProfile;

public class AliyunSmsSender {
	
	public static String confTemplate(){
		JSONObject jso = new JSONObject();
		jso.put("accessKey", "xxx");
		jso.put("accessSecret", "xxx");
		jso.put("signCode", "SC_xxx");
		jso.put("templateCode", "TC_xxx");
		jso.put("rowStart", "2");
		jso.put("rowEnd", "3");
		jso.put("recNum", "A");
		jso.put("param-1", "B");
		jso.put("param-2", "C");
		
		return JSON.toJSONString(jso, true);
	}
	
	/**
	 * 
	 * @param accessKeyId
	 * @param accessKeySecret
	 * @param sign        控制台创建的签名名称
	 * @param template  控制台创建的模板CODE
	 * @param parms
	 * @param recnum   接收号码
	 * @throws ClientException 
	 */
	public static String  send(String accessKeyId, String accessKeySecret, String sign, String template, String parms, String recnum) throws ClientException {
        //产品名称:云通信短信API产品,开发者无需替换
        String product = "Dysmsapi";
        //产品域名,开发者无需替换
        String domain = "dysmsapi.aliyuncs.com";

        //可自助调整超时时间
        System.setProperty("sun.net.client.defaultConnectTimeout", "10000");
        System.setProperty("sun.net.client.defaultReadTimeout", "10000");

        //初始化acsClient,暂不支持region化
        IClientProfile profile = DefaultProfile.getProfile("cn-hangzhou", accessKeyId, accessKeySecret);
        DefaultProfile.addEndpoint("cn-hangzhou", "cn-hangzhou", product, domain);
        IAcsClient acsClient = new DefaultAcsClient(profile);

        //组装请求对象-具体描述见控制台-文档部分内容
        SendSmsRequest request = new SendSmsRequest();
        //必填:待发送手机号
        request.setPhoneNumbers(recnum);
        //必填:短信签名-可在短信控制台中找到
        request.setSignName(sign);
        //必填:短信模板-可在短信控制台中找到
        request.setTemplateCode(template);
        //可选:模板中的变量替换JSON串,如模板内容为"亲爱的${name},您的验证码为${code}"时,此处的值为
        request.setTemplateParam(parms);

        //选填-上行短信扩展码(无特殊需求用户请忽略此字段)
        //request.setSmsUpExtendCode("90997");

        //可选:outId为提供给业务方扩展字段,最终在短信回执消息中将此值带回给调用者
        //request.setOutId("yourOutId");

        //hint 此处可能会抛出异常，注意catch
        SendSmsResponse sendSmsResponse = acsClient.getAcsResponse(request);

        return sendSmsResponse.getRequestId();
    }

}
