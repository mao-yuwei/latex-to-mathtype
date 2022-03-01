using org.apache.rocketmq.client.consumer;
using org.apache.rocketmq.client.consumer.listener;
using org.apache.rocketmq.client.producer;
using org.apache.rocketmq.common.consumer;
using org.apache.rocketmq.common.message;
using System;
using System.Text;
using System.Web.Script.Serialization;
using ConvertEquations.Comm;
using LatexToMathType.Models;
using System.Threading;
using System.Text.RegularExpressions;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;

namespace ConvertEquations
{
    class Program
    {
        public static object path;
        [STAThread]
        static void Main(string[] args)
        {

            Console.WriteLine("---------------------------------------------------------------------------------");
            Console.WriteLine("开始接收消息");
            Console.WriteLine("---------------------------------------------------------------------------------");

            OLE_WMF.InitWord();

            /*string latex = latexToMathType("{\"latex\":\"$2^a$\"}");
            Console.WriteLine(latex);*/

            #region 消费者
            DefaultMQPushConsumer consumer = new DefaultMQPushConsumer("latex-mathtype-request-group");
            consumer.setConsumeFromWhere(ConsumeFromWhere.CONSUME_FROM_FIRST_OFFSET);
            consumer.setNamesrvAddr("localhost:9876");
            ////Go_Ticket_WuLang_Test toptic名字，* 表示不过滤tag，如果过滤，可使用 ||划分
            consumer.subscribe("latex-to-mathtype-request-topic", "latex-to-mathtype-tag");
            consumer.registerMessageListener(new TestListener());
            consumer.setConsumeThreadMax(1);
            consumer.setConsumeThreadMin(1);
            consumer.setPullBatchSize(200);
            consumer.start();
            #endregion
        }


        public static string latexToMathType(string latexStr)
        {
            try
            {
                JavaScriptSerializer js = new JavaScriptSerializer();
                RequestParam latexMap = js.Deserialize<RequestParam>(latexStr);
                String mml = latexMap.mml;
                MathTypeModel mathType = new MathTypeModel();
                if (mml != null && !mml.Equals(""))
                {
                    Console.WriteLine("mml:" + latexMap.latex);
                    mathType = OLE_WMF.GetOLEAndWMFFromOneWordMML(mml);
                    mathType.type = "2";
                }
                else
                {
                    String latex = latexMap.latex;
                    String lLatex = latexMap.lLatex;
                    if (lLatex != null && !lLatex.Equals(""))
                    {
                        latex = lLatex;
                    }
                    latex = CleanLatex(latex);
                    if (InvalidLatex(latex))
                    {
                        Console.WriteLine("无效公式：" + latex);
                        return null;
                    }
                    mathType = OLE_WMF.GetOLEAndWMFFromOneWord(latex);
                }
                if (mathType == null) { return null; }
                mathType.lQid = latexMap.lQid;
                mathType.latex = latexMap.latex;
                string jsonData = js.Serialize(mathType);
                if (mathType.ole == null || mathType.wmf == null) { return null; }
                return jsonData;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Thread.Sleep(500);
            }
            return null;
        }



        /*  public class TestListener : MessageListenerOrderly
          {
              [STAThread]
              public ConsumeOrderlyStatus consumeMessage(java.util.List list, ConsumeOrderlyContext coc)
              {
                  for (int i = 0; i < list.size(); i++)
                  {
                      var msg = list.get(i) as Message;
                      byte[] body = msg.getBody();
                      var str = Encoding.UTF8.GetString(body);
                      Console.WriteLine("收到：" + str);
                      try
                      {
                          JavaScriptSerializer js = new JavaScriptSerializer();
                          RequestParam latexMap = js.Deserialize<RequestParam>(str);
                          String latex = latexMap.latex;
                          latex = CleanLatex(latex);
                          if (InvalidLatex(latex))
                          {
                              Console.WriteLine("无效公式：" + latex);
                              return ConsumeOrderlyStatus.SUCCESS;
                          }
                          // MathTypeModel mathType = OLE_WMF.GetOLEAndWMF(latex);
                          MathTypeModel mathType = OLE_WMF.GetOLEAndWMFFromWord(latex);
                          if (mathType == null) { return ConsumeOrderlyStatus.ROLLBACK; }

                          mathType.lQid = latexMap.lQid;
                          mathType.latex = latexMap.latex;
                          string jsonData = js.Serialize(mathType);
                          if (mathType.ole == null || mathType.wmf == null) { return ConsumeOrderlyStatus.ROLLBACK; }

                          ProducterSendMessage(jsonData);
                      }
                      catch (Exception e)
                      {
                          Console.WriteLine(e.Message);
                          Thread.Sleep(500);
                          return ConsumeOrderlyStatus.SUCCESS;
                      }
                      return ConsumeOrderlyStatus.SUCCESS;
                  }
                  return ConsumeOrderlyStatus.SUCCESS;
              }
          }*/

        public class TestListener : MessageListenerConcurrently
        {
            /// <summary>
            /// 消费功能，消费失败返回ConsumeConcurrentlyStatus.RECONSUME_LATER，成功返回CONSUME_SUCCESS
            /// </summary>
            /// <param name="list">虽然是个list，但是list的size是一个</param>
            /// <param name="ccc">上下文，对一些参数做设置，例如</param>
            /// <returns>结果</returns>
            [STAThread]
            public ConsumeConcurrentlyStatus consumeMessage(java.util.List list, ConsumeConcurrentlyContext ccc)
            {
                for (int i = 0; i < list.size(); i++)
                {
                    var msg = list.get(i) as Message;
                    byte[] body = msg.getBody();
                    var str = Encoding.UTF8.GetString(body);
                    Console.WriteLine("收到：" + str);
                    try
                    {
                        JavaScriptSerializer js = new JavaScriptSerializer();
                        RequestParam latexMap = js.Deserialize<RequestParam>(str);
                        String mml = latexMap.mml;
                        MathTypeModel mathType = new MathTypeModel();
                        if (mml != null && !mml.Equals(""))
                        {
                            Console.WriteLine("mml:" + latexMap.latex);
                            mathType = OLE_WMF.GetOLEAndWMFFromOneWordMML(mml);
                            mathType.type = "2";
                        }
                        else
                        {
                            String latex = latexMap.latex;
                            String lLatex = latexMap.lLatex;
                            if (lLatex != null && !lLatex.Equals(""))
                            {
                                latex = lLatex;
                            }
                            latex = CleanLatex(latex);
                            if (InvalidLatex(latex))
                            {
                                Console.WriteLine("无效公式：" + latex);
                                return ConsumeConcurrentlyStatus.CONSUME_SUCCESS;
                            }
                            //MathTypeModel mathType = OLE_WMF.GetOLEAndWMF(latex);
                            //MathTypeModel mathType = OLE_WMF.GetOLEAndWMFFromWord(latex);
                            mathType = OLE_WMF.GetOLEAndWMFFromOneWord(latex);
                        }
                        if (mathType == null) { return ConsumeConcurrentlyStatus.CONSUME_SUCCESS; }
                        //if (mathType == null) { return ConsumeConcurrentlyStatus.CONSUME_SUCCESS; }
                        mathType.lQid = latexMap.lQid;
                        mathType.latex = latexMap.latex;
                        string jsonData = js.Serialize(mathType);
                        if (mathType.ole == null || mathType.wmf == null) { return ConsumeConcurrentlyStatus.CONSUME_SUCCESS; }
                        //if (mathType.ole == null || mathType.wmf == null) { return ConsumeConcurrentlyStatus.CONSUME_SUCCESS; }
                        ProducterSendMessage(jsonData);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                        Thread.Sleep(500);
                        //return ConsumeOrderlyStatus.SUSPEND_CURRENT_QUEUE_A_MOMENT;
                        return ConsumeConcurrentlyStatus.CONSUME_SUCCESS;
                    }

                }
                return ConsumeConcurrentlyStatus.CONSUME_SUCCESS;
            }

        }

        public static Boolean InvalidLatex(String latex)
        {
            Regex r = new Regex("[^\\x00-\\xff]");
            Match m = r.Match(latex);
            if (m.Success) { return true; }
            Regex r1 = new Regex("\\$[ \\\\]*\\$");
            Match m1 = r1.Match(latex);
            if (m1.Success) { return true; }
            Regex r2 = new Regex("&.*?;");
            Match m2 = r2.Match(latex);
            if (m2.Success) { return true; }
            if (latex.Contains("⊄")) { return true;}
            if (latex.Contains("\\not ")) { return true;}
            if (latex.Contains("¢")) { return true;}
            if (latex.Contains("×")) { return true;}
            if (latex.Contains("boldsymbol")) { return true;}
            return false;
        }

        public static String CleanLatex(String latex)
        {
            #region
            latex = latex.Replace("&nbsp;", " ");
            latex = latex.Replace("&lt;", "<");
            latex = latex.Replace("&gt;", ">");
            latex = latex.Replace("&amp;", "&");
            latex = latex.Replace("&quot;", "\"");
            latex = latex.Replace("&apos;", "\'");
            latex = latex.Replace("&cent;", "¢");
            latex = latex.Replace("&pound;", "£");
            latex = latex.Replace("&yen;", "¥");
            latex = latex.Replace("&euro;", "€");
            latex = latex.Replace("&sect;", "§");
            latex = latex.Replace("&copy;", "©");
            latex = latex.Replace("&reg;", "®");
            latex = latex.Replace("&trade;", "™");
            latex = latex.Replace("&times;", "\\times ");
            latex = latex.Replace("&divide;", "÷");
            latex = latex.Replace("&ensp;", " ");
            latex = latex.Replace("&emsp;", " ");
            latex = latex.Replace("&Alpha;", "Α");
            latex = latex.Replace("&Gamma;", "Γ");
            latex = latex.Replace("&Epsilon;", "Ε");
            latex = latex.Replace("&Eta;", "Η");
            latex = latex.Replace("&Iota;", "Ι");
            latex = latex.Replace("&Lambda;", "Λ");
            latex = latex.Replace("&Nu;", "Ν");
            latex = latex.Replace("&Omicron;", "Ο");
            latex = latex.Replace("&Rho;", "Ρ");
            latex = latex.Replace("&Tau;", "Τ");
            latex = latex.Replace("&Phi;", "Φ");
            latex = latex.Replace("&Psi;", "Ψ");
            latex = latex.Replace("&alpha;", "α");
            latex = latex.Replace("&gamma;", "γ");
            latex = latex.Replace("&epsilon;", "ε");
            latex = latex.Replace("&eta;", "η");
            latex = latex.Replace("&iota;", "ι");
            latex = latex.Replace("&lambda;", "λ");
            latex = latex.Replace("&nu;", "ν");
            latex = latex.Replace("&omicron;", "ο");
            latex = latex.Replace("&rho;", "ρ");
            latex = latex.Replace("&sigma;", "σ");
            latex = latex.Replace("&upsilon;", "υ");
            latex = latex.Replace("&chi;", "χ");
            latex = latex.Replace("&omega;", "ω");
            latex = latex.Replace("&upsih;", "ϒ");
            latex = latex.Replace("&bull;", "•");
            latex = latex.Replace("&prime;", "′");
            latex = latex.Replace("&oline;", "‾");
            latex = latex.Replace("&weierp;", "℘");
            latex = latex.Replace("&real;", "ℜ");
            latex = latex.Replace("&alefsym;", "ℵ");
            latex = latex.Replace("&uarr;", "↑");
            latex = latex.Replace("&darr;", "↓");
            latex = latex.Replace("&crarr;", "↵");
            latex = latex.Replace("&uArr;", "⇑");
            latex = latex.Replace("&dArr;", "⇓");
            latex = latex.Replace("&forall;", "∀");
            latex = latex.Replace("&exist;", "∃");
            latex = latex.Replace("&nabla;", "∇");
            latex = latex.Replace("&notin;", "∉");
            latex = latex.Replace("&prod;", "∏");
            latex = latex.Replace("&minus;", "−");
            latex = latex.Replace("&radic;", "√");
            latex = latex.Replace("&infin;", "∞");
            latex = latex.Replace("&and;", "⊥");
            latex = latex.Replace("&cap;", "∩");
            latex = latex.Replace("&int;", "∫");
            latex = latex.Replace("&sim;", "∼");
            latex = latex.Replace("&asymp;", "≅");
            latex = latex.Replace("&equiv;", "≡");
            latex = latex.Replace("&ge;", "≥");
            latex = latex.Replace("&sup;", "⊃");
            latex = latex.Replace("&sube;", "⊆");
            latex = latex.Replace("&oplus;", "⊕");
            latex = latex.Replace("&perp;", "⊥");
            latex = latex.Replace("&lceil;", "⌈");
            latex = latex.Replace("&lfloor;", "⌊");
            latex = latex.Replace("&loz;", "◊");
            latex = latex.Replace("&clubs;", "♣");
            latex = latex.Replace("&diams;", "♦");
            latex = latex.Replace("&iexcl;", "¡");
            latex = latex.Replace("&laquo;", "«");
            latex = latex.Replace("&shy;", "­");
            latex = latex.Replace("&macr;", "¯");
            latex = latex.Replace("&plusmn;", "±");
            latex = latex.Replace("&sup3;", "³");
            latex = latex.Replace("&micro;", "µ");
            latex = latex.Replace("&Beta;", "Β");
            latex = latex.Replace("&Delta;", "Δ");
            latex = latex.Replace("&Zeta;", "Ζ");
            latex = latex.Replace("&Theta;", "Θ");
            latex = latex.Replace("&Kappa;", "Κ");
            latex = latex.Replace("&Mu;", "Μ");
            latex = latex.Replace("&Xi;", "Ξ");
            latex = latex.Replace("&Pi;", "Π");
            latex = latex.Replace("&Sigma;", "Σ");
            latex = latex.Replace("&Upsilon;", "Υ");
            latex = latex.Replace("&Chi;", "Χ");
            latex = latex.Replace("&Omega;", "Ω");
            latex = latex.Replace("&beta;", "β");
            latex = latex.Replace("&delta;", "δ");
            latex = latex.Replace("&zeta;", "ζ");
            latex = latex.Replace("&theta;", "θ");
            latex = latex.Replace("&kappa;", "κ");
            latex = latex.Replace("&mu;", "μ");
            latex = latex.Replace("&xi;", "ξ");
            latex = latex.Replace("&pi;", "π");
            latex = latex.Replace("&sigmaf;", "ς");
            latex = latex.Replace("&tau;", "τ");
            latex = latex.Replace("&phi;", "φ");
            latex = latex.Replace("&psi;", "ψ");
            latex = latex.Replace("&thetasym;", "ϑ");
            latex = latex.Replace("&piv;", "ϖ");
            latex = latex.Replace("&hellip;", "…");
            latex = latex.Replace("&Prime;", "″");
            latex = latex.Replace("&frasl;", "⁄");
            latex = latex.Replace("&image;", "ℑ");
            latex = latex.Replace("&larr;", "←");
            latex = latex.Replace("&rarr;", "→");
            latex = latex.Replace("&harr;", "↔");
            latex = latex.Replace("&lArr;", "⇐");
            latex = latex.Replace("&rArr;", "⇒");
            latex = latex.Replace("&hArr;", "⇔");
            latex = latex.Replace("&part;", "∂");
            latex = latex.Replace("&empty;", "∅");
            latex = latex.Replace("&isin;", "∈");
            latex = latex.Replace("&ni;", "∋");
            latex = latex.Replace("&sum;", "−");
            latex = latex.Replace("&lowast;", "∗");
            latex = latex.Replace("&prop;", "∝");
            latex = latex.Replace("&ang;", "∠");
            latex = latex.Replace("&or;", "⊦");
            latex = latex.Replace("&cup;", "∪");
            latex = latex.Replace("&there4;", "∴");
            latex = latex.Replace("&cong;", "≅");
            latex = latex.Replace("&ne;", "≠");
            latex = latex.Replace("&le;", "≤");
            latex = latex.Replace("&sub;", "⊂");
            latex = latex.Replace("&nsub;", "⊄");
            latex = latex.Replace("&supe;", "⊇");
            latex = latex.Replace("&otimes;", "⊗");
            latex = latex.Replace("&sdot;", "⋅");
            latex = latex.Replace("&rceil;", "⌉");
            latex = latex.Replace("&rfloor;", "⌋");
            latex = latex.Replace("&spades;", "♠");
            latex = latex.Replace("&hearts;", "♥");
            latex = latex.Replace("&curren;", "¤");
            latex = latex.Replace("&brvbar;", "¦");
            latex = latex.Replace("&uml;", "¨");
            latex = latex.Replace("&ordf;", "ª");
            latex = latex.Replace("&not;", "¬");
            latex = latex.Replace("&deg;", "°");
            latex = latex.Replace("&sup2;", "²");
            latex = latex.Replace("&acute;", "´");
            latex = latex.Replace("&middot;", "·");
            latex = latex.Replace("&oslash;", "ø");
            latex = latex.Replace("&aacute;", "á");
            latex = latex.Replace("&#160;", "");
            latex = latex.Replace("&#60;", "<");
            latex = latex.Replace("&#62;", ">");
            latex = latex.Replace("&#38;", "&");
            latex = latex.Replace("&#34;", "\"");
            latex = latex.Replace("&#39;", "\'");
            latex = latex.Replace("&#162;", "¢");
            latex = latex.Replace("&#163;", "£");
            latex = latex.Replace("&#165;", "¥");
            latex = latex.Replace("&#8364;", "€");
            latex = latex.Replace("&#167;", "§");
            latex = latex.Replace("&#169;", "©");
            latex = latex.Replace("&#174;", "®");
            latex = latex.Replace("&#8482;", "™");
            latex = latex.Replace("&#215;", "\\times");
            latex = latex.Replace("&#247;", "÷");
            latex = latex.Replace("&#160;", "");
            latex = latex.Replace("&#160;", "");
            latex = latex.Replace("&#913;", "Α");
            latex = latex.Replace("&#915;", "Γ");
            latex = latex.Replace("&#917;", "Ε");
            latex = latex.Replace("&#919;", "Η");
            latex = latex.Replace("&#921;", "Ι");
            latex = latex.Replace("&#923;", "Λ");
            latex = latex.Replace("&#925;", "Ν");
            latex = latex.Replace("&#927;", "Ο");
            latex = latex.Replace("&#929;", "Ρ");
            latex = latex.Replace("&#932;", "Τ");
            latex = latex.Replace("&#934;", "Φ");
            latex = latex.Replace("&#936;", "Ψ");
            latex = latex.Replace("&#945;", "α");
            latex = latex.Replace("&#947;", "γ");
            latex = latex.Replace("&#949;", "ε");
            latex = latex.Replace("&#951;", "η");
            latex = latex.Replace("&#953;", "ι");
            latex = latex.Replace("&#955;", "λ");
            latex = latex.Replace("&#957;", "ν");
            latex = latex.Replace("&#959;", "ο");
            latex = latex.Replace("&#961;", "ρ");
            latex = latex.Replace("&#963;", "σ");
            latex = latex.Replace("&#965;", "υ");
            latex = latex.Replace("&#967;", "χ");
            latex = latex.Replace("&#969;", "ω");
            latex = latex.Replace("&#978;", "ϒ");
            latex = latex.Replace("&#8226;", "•");
            latex = latex.Replace("&#8242;", "′");
            latex = latex.Replace("&#8254;", "‾");
            latex = latex.Replace("&#8472;", "℘");
            latex = latex.Replace("&#8476;", "ℜ");
            latex = latex.Replace("&#8501;", "ℵ");
            latex = latex.Replace("&#8593;", "↑");
            latex = latex.Replace("&#8595;", "↓");
            latex = latex.Replace("&#8629;", "↵");
            latex = latex.Replace("&#8657;", "⇑");
            latex = latex.Replace("&#8659;", "⇓");
            latex = latex.Replace("&#8704;", "∀");
            latex = latex.Replace("&#8707;", "∃");
            latex = latex.Replace("&#8711;", "∇");
            latex = latex.Replace("&#8713;", "∉");
            latex = latex.Replace("&#8719;", "∏");
            latex = latex.Replace("&#8722;", "−");
            latex = latex.Replace("&#8730;", "√");
            latex = latex.Replace("&#8734;", "∞");
            latex = latex.Replace("&#8869;", "⊥");
            latex = latex.Replace("&#8745;", "∩");
            latex = latex.Replace("&#8747;", "∫");
            latex = latex.Replace("&#8764;", "∼");
            latex = latex.Replace("&#8773;", "≅");
            latex = latex.Replace("&#8801;", "≡");
            latex = latex.Replace("&#8805;", "≥");
            latex = latex.Replace("&#8835;", "⊃");
            latex = latex.Replace("&#8838;", "⊆");
            latex = latex.Replace("&#8853;", "⊕");
            latex = latex.Replace("&#8869;", "⊥");
            latex = latex.Replace("&#8968;", "⌈");
            latex = latex.Replace("&#8970;", "⌊");
            latex = latex.Replace("&#9674;", "◊");
            latex = latex.Replace("&#9827;", "♣");
            latex = latex.Replace("&#9830;", "♦");
            latex = latex.Replace("&#161;", "¡");
            latex = latex.Replace("&#171;", "«");
            latex = latex.Replace("&#173;", "­");
            latex = latex.Replace("&#175;", "¯");
            latex = latex.Replace("&#177;", "±");
            latex = latex.Replace("&#179;", "³");
            latex = latex.Replace("&#181;", "µ");
            latex = latex.Replace("&#914;", "Β");
            latex = latex.Replace("&#916;", "Δ");
            latex = latex.Replace("&#918;", "Ζ");
            latex = latex.Replace("&#920;", "Θ");
            latex = latex.Replace("&#922;", "Κ");
            latex = latex.Replace("&#924;", "Μ");
            latex = latex.Replace("&#926;", "Ξ");
            latex = latex.Replace("&#928;", "Π");
            latex = latex.Replace("&#931;", "Σ");
            latex = latex.Replace("&#933;", "Υ");
            latex = latex.Replace("&#935;", "Χ");
            latex = latex.Replace("&#937;", "Ω");
            latex = latex.Replace("&#946;", "β");
            latex = latex.Replace("&#948;", "δ");
            latex = latex.Replace("&#950;", "ζ");
            latex = latex.Replace("&#952;", "θ");
            latex = latex.Replace("&#954;", "κ");
            latex = latex.Replace("&#956;", "μ");
            latex = latex.Replace("&#958;", "ξ");
            latex = latex.Replace("&#960;", "π");
            latex = latex.Replace("&#962;", "ς");
            latex = latex.Replace("&#964;", "τ");
            latex = latex.Replace("&#966;", "φ");
            latex = latex.Replace("&#968;", "ψ");
            latex = latex.Replace("&#977;", "ϑ");
            latex = latex.Replace("&#982;", "ϖ");
            latex = latex.Replace("&#8230;", "…");
            latex = latex.Replace("&#8243;", "″");
            latex = latex.Replace("&#8260;", "⁄");
            latex = latex.Replace("&#8465;", "ℑ");
            latex = latex.Replace("&#8592;", "←");
            latex = latex.Replace("&#8594;", "→");
            latex = latex.Replace("&#8596;", "↔");
            latex = latex.Replace("&#8656;", "⇐");
            latex = latex.Replace("&#8658;", "⇒");
            latex = latex.Replace("&#8660;", "⇔");
            latex = latex.Replace("&#8706;", "∂");
            latex = latex.Replace("&#8709;", "∅");
            latex = latex.Replace("&#8712;", "∈");
            latex = latex.Replace("&#8715;", "∋");
            latex = latex.Replace("&#8722;", "−");
            latex = latex.Replace("&#8727;", "∗");
            latex = latex.Replace("&#8733;", "∝");
            latex = latex.Replace("&#8736;", "∠");
            latex = latex.Replace("&#8870;", "⊦");
            latex = latex.Replace("&#8746;", "∪");
            latex = latex.Replace("&#8756;", "∴");
            latex = latex.Replace("&#8773;", "≅");
            latex = latex.Replace("&#8800;", "≠");
            latex = latex.Replace("&#8804;", "≤");
            latex = latex.Replace("&#8834;", "⊂");
            latex = latex.Replace("&#8836;", "⊄");
            latex = latex.Replace("&#8839;", "⊇");
            latex = latex.Replace("&#8855;", "⊗");
            latex = latex.Replace("&#8901;", "⋅");
            latex = latex.Replace("&#8969;", "⌉");
            latex = latex.Replace("&#8971;", "⌋");
            latex = latex.Replace("&#9824;", "♠");
            latex = latex.Replace("&#9829;", "♥");
            latex = latex.Replace("&#164;", "¤");
            latex = latex.Replace("&#166;", "¦");
            latex = latex.Replace("&#168;", "¨");
            latex = latex.Replace("&#170;", "ª");
            latex = latex.Replace("&#172;", "¬");
            latex = latex.Replace("&#176;", "°");
            latex = latex.Replace("&#178;", "²");
            latex = latex.Replace("&#180;", "´");

            latex = latex.Replace("θ", "\\theta ");
            latex = latex.Replace("π", "\\pi ");
            latex = latex.Replace("α", "\\alpha ");
            latex = latex.Replace("β", "\\beta ");
            latex = latex.Replace("＜", "<");
            latex = latex.Replace("＞", ">");
            latex = latex.Replace("∵", "\\because ");
            latex = latex.Replace("∴", "\\therefore ");
            latex = latex.Replace("&lt;", "<");
            latex = latex.Replace("&gt;", ">");
            latex = latex.Replace("&times;", "\\times ");
            latex = latex.Replace("&#215;", "\\times ");
            latex = latex.Replace("&amp;", "& ");
            latex = latex.Replace("×", "\\times ");
            latex = latex.Replace("&#60;", "<");
            latex = latex.Replace("&#62;", ">");
            latex = latex.Replace("&#38;", "&");
            latex = latex.Replace("∈", "\\in ");
            latex = latex.Replace("∉", "\\notin ");
            latex = latex.Replace("ρ", "\\rho ");
            latex = latex.Replace("≥", "\\geqslant ");
            latex = latex.Replace("≤", "\\leqslant ");
            latex = latex.Replace("：", ":");
            latex = latex.Replace("，", ",");
            latex = latex.Replace("…", "...");
            latex = latex.Replace("△", "\\triangle ");
            latex = latex.Replace("□", "\\square ");
            latex = latex.Replace("∥", "\\parallel ");
            latex = latex.Replace("∠", "\\angle ");
            latex = latex.Replace("⊥", "\\bot ");
            latex = latex.Replace("′", "'");
            latex = latex.Replace("⋅", "\\cdot ");
            latex = latex.Replace("≈", "\\thickapprox ");
            latex = latex.Replace("{{'}}", "'");
            latex = latex.Replace("&plusmn;", "\\pm ");
            latex = latex.Replace("&#177;", "\\pm ");
            latex = latex.Replace("∞", "\\infty ");
            latex = latex.Replace("∪", "\\cup ");
            latex = latex.Replace("φ", "\\varphi ");

            latex = latex.Replace("α", "\\alpha ");
            latex = latex.Replace("η", "\\eta ");
            latex = latex.Replace("γ", "\\gamma ");
            latex = latex.Replace("δ", "\\delta ");
            latex = latex.Replace("φ", "\\phi ");
            latex = latex.Replace("σ", "\\sigma ");
            latex = latex.Replace("Σ", "\\Sigma ");
            latex = latex.Replace("φ", "\\varphi ");
            latex = latex.Replace("∂", "\\partial ");
            latex = latex.Replace("∞", "\\infty ");
            latex = latex.Replace("≠", "\\neq ");
            latex = latex.Replace("≠", "\\ne ");
            latex = latex.Replace("±", "\\pm ");
            latex = latex.Replace("≈", "\\approx ");
            latex = latex.Replace("β", "\\beta ");
            latex = latex.Replace("ε", "\\varepsilon ");
            latex = latex.Replace("ϵ", "\\epsilon ");
            latex = latex.Replace("μ", "\\mu ");
            latex = latex.Replace("ω", "\\omega ");
            latex = latex.Replace("π", "\\pi ");
            latex = latex.Replace("×", "\\times ");
            latex = latex.Replace("÷", "\\div ");
            latex = latex.Replace("∝", "\\propto ");
            latex = latex.Replace("∀", "\\forall ");
            latex = latex.Replace("∃", "\\exists ");
            latex = latex.Replace("∈", "\\in ");
            latex = latex.Replace("∋", "\\ni ");
            latex = latex.Replace("∪", "\\cup ");
            latex = latex.Replace("∩", "\\cap ");
            latex = latex.Replace("∴", "\\therefore ");
            latex = latex.Replace("∵", "\\because ");
            latex = latex.Replace("ξ", "\\xi ");
            latex = latex.Replace("ρ", "\\rho ");
            latex = latex.Replace("τ", "\\tau ");
            latex = latex.Replace("υ", "\\upsilon ");
            latex = latex.Replace("λ", "\\lambda ");
            latex = latex.Replace("Δ", "\\Delta ");
            latex = latex.Replace("Γ", "\\Gamma ");
            latex = latex.Replace("Π", "\\Pi ");
            latex = latex.Replace("←", "\\gets ");
            latex = latex.Replace("←", "\\leftarrow ");
            latex = latex.Replace("→", "\\rightarrow ");
            latex = latex.Replace("→", "\\to ");
            latex = latex.Replace("↑", "\\uparrow ");
            latex = latex.Replace("↓", "\\downarrow ");
            latex = latex.Replace("↔", "\\leftrightarrow ");
            latex = latex.Replace("↕", "\\updownarrow ");
            latex = latex.Replace("⇐", "\\Leftarrow ");
            latex = latex.Replace("⇒", "\\Rightarrow ");
            latex = latex.Replace("⇑", "\\Uparrow ");
            latex = latex.Replace("⇓", "\\Downarrow ");
            latex = latex.Replace("⇔", "\\Leftrightarrow ");
            latex = latex.Replace("⇕", "\\Updownarrow ");
            latex = latex.Replace("⟵", "\\longleftarrow ");
            latex = latex.Replace("⟶", "\\longrightarrow ");
            latex = latex.Replace("⟷", "\\longleftrightarrow ");
            latex = latex.Replace("⟸", "\\Longleftarrow ");
            latex = latex.Replace("⟹", "\\Longrightarrow ");
            latex = latex.Replace("⟺", "\\Longleftrightarrow ");
            latex = latex.Replace("↗", "\\nearrow ");
            latex = latex.Replace("↖", "\\nwarrow ");
            latex = latex.Replace("∠", "\\angle ");
            latex = latex.Replace("⊥", "\\bot ");
            latex = latex.Replace("∥", "\\parallel ");
            // latex = latex.Replace("△","\\triangle ");
            latex = latex.Replace("□", "\\Box ");
            latex = latex.Replace("□", "\\square ");
            latex = latex.Replace("◊", "\\Diamond ");
            latex = latex.Replace("∘", "\\circ ");
            latex = latex.Replace("∘", "\\omicron ");
            latex = latex.Replace("⊕", "\\oplus ");
            latex = latex.Replace("⊕", "\\bigoplus ");
            latex = latex.Replace("⊗", "\\otimes ");
            latex = latex.Replace("⊗", "\\bigotimes ");
            latex = latex.Replace("·", "\\bullet ");
            latex = latex.Replace("∙", "\\cdot ");
            latex = latex.Replace("♣", "\\clubsuit ");
            latex = latex.Replace("♠", "\\spadesuit ");
            latex = latex.Replace("♯", "\\sharp ");
            latex = latex.Replace("…", "\\ldots ");
            latex = latex.Replace("…", "\\dots ");
            latex = latex.Replace("'", "\\prime ");
            latex = latex.Replace("∧", "\\wedge ");
            latex = latex.Replace("∧", "\\land ");
            latex = latex.Replace("∧", "\\bigwedge ");
            latex = latex.Replace("∨", "\\lor ");
            latex = latex.Replace("∨", "\\vee ");
            latex = latex.Replace("∨", "\\bigvee ");
            latex = latex.Replace("¬", "\\lnot ");
            latex = latex.Replace("¬", "\\neg ");
            latex = latex.Replace("∖", "\\setminus ");
            latex = latex.Replace("∇", "\\nabla ");
            latex = latex.Replace("~", "\\sim ");
            latex = latex.Replace("⇀", "\\rightharpoonup ");
            latex = latex.Replace("⇁", "\\rightharpoondown ");
            latex = latex.Replace("↽", "\\leftharpoondown ");
            latex = latex.Replace("↼", "\\leftharpoonup ");
            latex = latex.Replace("↿", "\\upharpoonleft ");
            latex = latex.Replace("↾", "\\upharpoonright ");
            latex = latex.Replace("⇃", "\\downharpoonleft ");
            latex = latex.Replace("⇂", "\\downharpoonright ");
            latex = latex.Replace("%", "\\% ");
            latex = latex.Replace("℧", "\\mho ");
            latex = latex.Replace("ℏ", "\\hbar ");
            latex = latex.Replace("⋯", "\\cdots ");
            latex = latex.Replace("∫", "\\int ");
            latex = latex.Replace("∑", "\\sum ");
            latex = latex.Replace("∏", "\\prod ");
            latex = latex.Replace("∬", "\\iint ");
            latex = latex.Replace("∭", "\\iiint ");
            latex = latex.Replace("∩", "\\bigcap ");
            latex = latex.Replace("∪", "\\bigcup ");
            latex = latex.Replace("⟨", "\\langle ");
            latex = latex.Replace("⟩", "\\rangle ");
            latex = latex.Replace("‖", "\\lvert ");
            latex = latex.Replace("‖", "\\rvert ");
            latex = latex.Replace("∮", "\\oint ");
            latex = latex.Replace("∐", "\\coprod ");
            latex = latex.Replace("≥", "\\geq ");
            latex = latex.Replace("⊂", "\\subset ");
            latex = latex.Replace("⊃", "\\supset ");
            latex = latex.Replace(":", "\\colon ");
            latex = latex.Replace("&#8801;", "\\equiv ");
            latex = latex.Replace("~", "\\~ ");
            latex = latex.Replace("⩾", "\\geqslant ");
            latex = latex.Replace("∅", "\\varnothing ");
            latex = latex.Replace("∉", "\\notin ");
            latex = latex.Replace("⩽", "\\leqslant ");
            latex = latex.Replace("≅", "\\cong ");
            latex = latex.Replace("∁", "\\complement ");
            latex = latex.Replace("△", "\\bigtriangleup ");
            latex = latex.Replace("⊙", "\\odot ");
            latex = latex.Replace("∣", "\\mid ");
            latex = latex.Replace("≜", "\\triangleq ");
            latex = latex.Replace("Ω", "\\Omega ");
            latex = latex.Replace("⇌", "\\rightleftharpoons ");
            latex = latex.Replace("⇋", "\\leftrightharpoons ");
            latex = latex.Replace("⇌", "\\rightleftarrows ");
            latex = latex.Replace("⇋", "\\leftrightarrows ");
            latex = latex.Replace("~", "\\tilde ");
            latex = latex.Replace("△", "\\vartriangle ");
            latex = latex.Replace("Λ", "\\Lambda ");
            latex = latex.Replace("ν", "\\nu ");
            latex = latex.Replace("·", "\\centerdot ");
            latex = latex.Replace("Θ", "\\Theta ");
            latex = latex.Replace("≫", "\\gg ");
            latex = latex.Replace("≪", "\\ll ");
            latex = latex.Replace("◯ ", "\\bigcirc ");
            latex = latex.Replace("∽", "\\backsim ");
            latex = latex.Replace("ϑ", "\\vartheta ");
            latex = latex.Replace("▲", "\\blacktriangle ");
            latex = latex.Replace("√", "\\surd ");
            latex = latex.Replace("⊆", "\\subseteq ");
            latex = latex.Replace("⊇", "\\supseteq ");
            latex = latex.Replace("∅", "\\emptyset ");
            latex = latex.Replace("Γ", "\\varGamma ");
            latex = latex.Replace("◊", "\\lozenge ");
            latex = latex.Replace("⌒", "\\frown ");
            latex = latex.Replace("ζ", "\\zeta ");
            latex = latex.Replace("ψ", "\\psi ");
            latex = latex.Replace("ς", "\\varsigma ");
            latex = latex.Replace("█", "\\blacksquare ");
            latex = latex.Replace("⋇", "\\divideontimes ");
            latex = latex.Replace("⋆", "\\star ");
            latex = latex.Replace("⊊", "\\subsetneq ");
            latex = latex.Replace("¬", "\\urcorner ");
            latex = latex.Replace("▻", "\\triangleright ");
            latex = latex.Replace("◁", "\\triangleleft ");
            latex = latex.Replace("∇", "\\triangledown ");
            latex = latex.Replace("✓", "\\checkmark ");
            latex = latex.Replace("*", "\\ast ");
            latex = latex.Replace("◎", "\\circledcirc ");
            latex = latex.Replace("≈", "\\thickapprox ");
            latex = latex.Replace("≺", "\\prec ");
            latex = latex.Replace("⋮", "\\vdots ");
            latex = latex.Replace("⋰", "\\rddots ");
            latex = latex.Replace("⋱", "\\ddots ");
            latex = latex.Replace("☉", "\\bigodot ");
            latex = latex.Replace("◆", "\\blacklozenge ");
            latex = latex.Replace("◊", "\\diamond ");
            latex = latex.Replace("♢", "\\diamondsuit ");
            latex = latex.Replace("∆", "\\Delta ");
            latex = latex.Replace("↘", "\\searrow ");
            latex = latex.Replace("↙", "\\swarrow ");
            latex = latex.Replace("⊈", "\\nsubseteq ");
            latex = latex.Replace("⇏", "\\nRightarrow ");
            latex = latex.Replace("↚", "\\nleftarrow ");
            latex = latex.Replace("≻", "\\succ ");
            latex = latex.Replace("═", "=");


            latex = latex.Replace("ϕ", "\\varphi ");
            latex = latex.Replace("•", "\\bullet ");
            latex = latex.Replace("（", " ( ");
            latex = latex.Replace("）", " ) ");
            latex = latex.Replace("＋", "+");
            latex = latex.Replace("＝", "=");
            latex = latex.Replace("−", "-");
            latex = latex.Replace("丨", "|");
            latex = latex.Replace("；", ";");
            latex = latex.Replace("－", "-");
            latex = latex.Replace("\\%", "%");
            latex = latex.Replace("\\\\%", "%");
            latex = latex.Replace("﹪", "%");
            latex = latex.Replace("￢", "{}^\\neg ");
            latex = latex.Replace("^{°}", "^{\\circ}");
            latex = latex.Replace("°", "^\\circ ");
            latex = latex.Replace("&lt;", "<");
            latex = latex.Replace("&gt;", ">");
            latex = latex.Replace("\\varDelta ", "\\Delta");
            latex = latex.Replace("\\lt ", "<");
            latex = latex.Replace("\\gt ", ">");
            latex = latex.Replace("\\lt\\", "<\\");
            latex = latex.Replace("\\gt\\", ">\\");
            latex = latex.Replace("\\lt{", "<{");
            latex = latex.Replace("\\gt{", ">{");
            latex = latex.Replace("\\lt}", "<}");
            latex = latex.Replace("\\gt}", ">}");
            latex = latex.Replace("\\lt=", "<=");
            latex = latex.Replace("\\gt=", ">=");
            latex = latex.Replace("\\$", "$");
            latex = latex.Replace("\\overset{^}", "\\hat ");
            latex = latex.Replace("\\overset{\\Large\\backsim }{=}", "\\cong ");
            latex = latex.Replace("\\Large", " ");
            latex = latex.Replace("\\sqrt[^3]", "\\sqrt[3]");
            latex = latex.Replace("’", "\\prime ");
            latex = latex.Replace("′", "\\prime ");
            latex = latex.Replace("", "");
            latex = latex.Replace("\u0082", " ");
            latex = latex.Replace("", " ");
            latex = latex.Replace("", " ");
            latex = latex.Replace("", " ");
            latex = latex.Replace("", " ");
            latex = latex.Replace("", " ");
            latex = latex.Replace("\\perp", "\\bot");
            latex = latex.Replace("⨂", "\\otimes ");


            latex = latex.Replace("\\\\colon /\\!/", "//");
            latex = latex.Replace("{^", "{{}^");
            latex = latex.Replace("{_", "{{}_");
            latex = latex.Replace("{ ^", "{{}^");
            latex = latex.Replace("{ _", "{{}_");
            latex = latex.Replace("( _", "( {}_");
            latex = latex.Replace("(_", "({}_");
            latex = latex.Replace("( ^", "( {}^");
            latex = latex.Replace("(^", "({}^");
            latex = latex.Replace("$\\\\=", "$=");
            latex = latex.Replace("\\lvert", "|");
            latex = latex.Replace("\\gt", ">");
            latex = latex.Replace("\\lt", "<");
            latex = latex.Replace("\\overset{\\hat }", "\\hat");
            latex = latex.Replace("\\rvert", "|");
            latex = latex.Replace("\\small ", "");
            latex = latex.Replace("\\&", "&");
            latex = latex.Replace("\\(-\\)", "-");
            latex = latex.Replace("\\:/\\!/", "//");
            latex = latex.Replace("\\\\colon / \\!/", "//");
            latex = latex.Replace("\\cot ", "cot");
            latex = latex.Replace("\\sec ", "sec");
            latex = latex.Replace("\"\"", "");
            latex = latex.Replace("\\sqrt[^{3}]", "\\sqrt[3]");
            latex = latex.Replace("\\varDelta", "\\Delta ");
            latex = latex.Replace("×", "\\times ");
            latex = latex.Replace("\\varGamma", "\\Gamma ");
            latex = latex.Replace("\\verb|//|", "//");
            latex = latex.Replace("´", "^\\prime ");
            latex = latex.Replace("º", "^{\\circ }");
            latex = latex.Replace("\\\\\\sec", "sec");
            latex = latex.Replace("$\\\\ =", "$=");
            latex = latex.Replace("∆", "\\Delta");
            latex = latex.Replace("\\overset{¯}", "\\overline");
            latex = latex.Replace("\\\\ \\end", "\\end");
            latex = latex.Replace("\\\\\\end", "\\end");
            latex = latex.Replace("\\\\\\sec", "sec");
            latex = latex.Replace("& amp;", "&");
            latex = latex.Replace("{\\kern 1pt}", "");
            latex = latex.Replace("\\mathop", "");
            latex = latex.Replace("\\%", "%");
            latex = latex.Replace("\u0082", " ");
            latex = latex.Replace("\uEF05", " ");
            latex = latex.Replace("\uF070", " ");
            latex = latex.Replace("\uE004", " ");
            latex = latex.Replace("\uF05C", " ");
            latex = latex.Replace("\uE003", " ");
            latex = latex.Replace("\u0083", " ");
            latex = latex.Replace("\u0081", " ");
            latex = latex.Replace("\uE00A", " ");
            latex = latex.Replace("\uF03D", " ");
            latex = latex.Replace("\uF020", " ");
            latex = latex.Replace("\uE584", " ");
            latex = latex.Replace("\u0084", " ");
            latex = latex.Replace("\uF0B0", " ");
            latex = latex.Replace("\\perp", "\\bot");
            latex = latex.Replace("⊥", "\\bot ");
            latex = latex.Replace("⨂", "\\otimes ");
            latex = latex.Replace("ɛ", "\\varepsilon ");
            latex = latex.Replace("∰", "\\volintegral ");
            latex = latex.Replace("∯", "\\surfintegral ");
            latex = latex.Replace("\\:", " ");
            latex = latex.Replace("\\;", " ");
            latex = latex.Replace("ˈ", "'");
            latex = latex.Replace("ʹ", "'");
            latex = latex.Replace("\\large", " ");
            latex = latex.Replace("\uF03E", " ");
            latex = latex.Replace("\\verb |//|", "//");
            latex = latex.Replace("\\;\\rm", "");
            latex = latex.Replace("❈", "※");
            latex = latex.Replace("\\vartriangle", "\\triangle");
            latex = latex.Replace("\\Longleftrightarrow", "\\leftrightarrow ");
            latex = latex.Replace("\\text{-}", "-");
            latex = latex.Replace("^{\\circ }","\\circ");
            latex = latex.Replace("^{\\circ}","\\circ");
            latex = latex.Replace("^\\circ", "\\circ");

            #endregion
            latex = TranLatex(latex);
            return latex;
        }

        public static String TranLatex(String latex)
        {
            if (latex.Contains("\\end{cases}"))
            {
                Regex r = new Regex("\\\\\\\\[ ]*\\\\end\\{cases");
                Match m = r.Match(latex);
                if (m.Success) { return latex; }
                latex = latex.Replace("\\end{cases}", "\\\\ \\end{cases}");
            }
            return latex;
        }



        static void Main1(string[] args)
        {
            OLE_WMF.InitWord();
            String mml = "<math>     <msqrt>         <mtext>于否</mtext>     </msqrt> </math>";
            OLE_WMF.GetOLEAndWMFFromOneWordMML(mml);
            String latex = "$\\odot O$";
            latex = CleanLatex(latex);
            // MathTypeModel mathType = OLE_WMF.GetOLEAndWMF(latex);
            MathTypeModel mathType = OLE_WMF.GetOLEAndWMFFromWord(latex);



            // byte[] wmf = Convert.FromBase64String(mathType.wmf);

            //String bs = "183GmgAAAAAAAGACIAICCQAAAABTXgEACQAAA6QBAAAHAJAAAAAAAAUAAAACAQEAAAAFAAAAAQL///8ABQAAAC4BGQAAAAUAAAALAgAAAAAFAAAADAIgAmACCwAAACYGDwAMAE1hdGhUeXBlAAAwABIAAAAmBg8AGgD/////AAAQAAAAwP///7r///8gAgAA2gEAAAgAAAD6AgAAAAAAAAAAAAIEAAAALQEAAAUAAAAUAl0BQwAFAAAAEwJIAWsABQAAABMC3gHHAAUAAAATAk8ALQEFAAAAEwJPAAoCBwAAAPwCAAAAAAACAAAEAAAALQEBAAgAAAD6AgUAAQAAAAAAAAAEAAAALQECABoAAAAkAwsAQABYAXUAOwHIALYBJgFGAAoCRgAKAlkANAFZANEA3gG+AN4BYQBVAUYAYwEEAAAALQEAAAUAAAAJAgAAAAIFAAAAFALAAT4BHAAAAPsCgP4AAAAAAACQAQAAAAAAAgAQVGltZXMgTmV3IFJvbWFuAAAACgAAAAAA8obWdEAAAAAEAAAALQEDAAkAAAAyCgAAAAABAAAAMgAAA5AAAAAmBg8AFQFBcHBzTUZDQwEA7gAAAO4AAABEZXNpZ24gU2NpZW5jZSwgSW5jLgAFAQAGCURTTVQ2AAETV2luQWxsQmFzaWNDb2RlUGFnZXMAEQVUaW1lcyBOZXcgUm9tYW4AEQNTeW1ib2wAEQVDb3VyaWVyIE5ldwARBE1UIEV4dHJhABNXaW5BbGxDb2RlUGFnZXMAEQbLzszlABIACCEvRY9EL0FQ9BAPR19BUPIfHkFQ9BUPQQD0RfQl9I9CX0EA9BAPQ19BAPSPRfQqX0j0j0EA9BAPQPSPQX9I9BAPQSpfRF9F9F9F9F9BDwwBAAEAAQICAgIAAgABAQEAAwABAAQABQAKAQADAAoAAAEAAgCIMgAACwEBAAAAAQoAAAAmBg8ACgD/////AQAAAAAACAAAAPoCAAAAAAAAAAAAAAQAAAAtAQQABwAAAPwCAAAAAAAAAAAEAAAALQEFABwAAAD7AhAABwAAAAAAvAIAAACGAQICIlN5c3RlbQAASACKAAAACgCFN2bVSACKAP////8Q1BkABAAAAC0BBgAEAAAA8AEDAAMAAAAAAA==";
            //wmf= Convert.FromBase64String(bs);
            // Image image = Image.FromStream(new MemoryStream(wmf));
            //image.Save("E:\\image\\image1.wmf", ImageFormat.Emf);


            /*  Metafile mf = new Metafile(new MemoryStream(wmf));

            //  Graphics gs = Graphics.FromImage(mf);
             // Metafile wf = new Metafile(@"E:\\image\\image1.wmf",gs.GetHdc());
              Metafile wf = new Metafile(@"D:\\image1.wmf",gs.GetHdc());
              Graphics g = Graphics.FromImage(wf);
              g.Dispose();
              mf.Dispose();
              wf.Dispose();
              //  mf.Save("E:\\image\\image1.wmf");

  */
            //Byte[] bb = ms.ToArray();

            //output("E:\\image\\image1.wmf", bb);
            JavaScriptSerializer js = new JavaScriptSerializer();
            string jsonData = js.Serialize(mathType);
            Console.WriteLine(jsonData);
            Console.Read();

        }

        public static void output(string path, byte[] fileByte)
        {

            FileStream fs = new FileStream(path, FileMode.Create);

            fs.Write(fileByte, 0, fileByte.Length);

            fs.Close();


        }


        public static Boolean ProducterSendMessage(String message)
        {


            #region 生产消息
            DefaultMQProducer producer = null;
            try
            {
                ////wulangtes  是分组名字，用来区分生产者
                producer = new DefaultMQProducer("latex-mathtype-response-group");

                ////服务器ip
                producer.setNamesrvAddr("localhost:9876");
                producer.start();

                ////Go_Ticket_WuLang_Test 为 toptic 名字，taga是比toptic更为精确地内容划分，RocketMQ会重试是内容
                Message msg = new Message("latex-to-mathtype-response-topic", "latex-to-mathtype-tag", Encoding.UTF8.GetBytes(message));
                SendResult sendResult = producer.send(msg);
                Console.WriteLine(sendResult.getSendStatus());

            }
            finally
            {
                producer.shutdown();
            }
            return true;

            #endregion

        }

    }
}
