说明：
　　凡于2008.1.1后购买的语音报价器，动态库改为TdBjq.dll和EasyD12.dll，
对于使用以前生产的报价器，程序不用改动，只需将TdBjq.dll、EasyD12.dll和
setting.ini文件拷入HIS系统调用目录即可，setting.ini文件中参数与以前设置
相同，但ComMode=1时代表使用串口的语音报价器，ComMode=0时代表使用标准USB口
的语音报价器。特此说明。

    使用USB接口的报价器,请先安装USB驱动程序。

    尊敬的用户，您在使用或接入中有何问题，请致电021-66080281，我们竭诚
为您服务。

以下新增协议:

'请您到预检处预检'                    ---协议字符为：v
'请问您孩子的姓名'                    ---协议字符为：w
'请出示您的挂号票'                    ---协议字符为：x
'请您到放射科，B超室，CT室划价!'      ---协议字符为：y
'请您到放射科划价'                    ---协议字符为：@
'请您到B超室划价'                     ---协议字符为：#
'请您到CT室划价'                      ---协议字符为：%
'请您到挂号处挂号'                    ---协议字符为：z


.请输入密码                           ---协议字符为：*
.请出示重症病历                       ---协议字符为：-
.本次费用总额                         ---协议字符为：K
.医保负担                             ---协议字符为：L
.医保卡支付                           ---协议字符为：M
.需支付现金                           ---协议字符为：N
.医保卡余额                           ---协议字符为：O
.公费负担                             ---协议字符为：Q
.医院优惠                             ---协议字符为：R
.个人支付                             ---协议字符为：S
.请收好您的医保卡                     ---协议字符为：T
.请出示医保卡                         ---协议字符为：U
.请出示公费医疗证                     ---协议字符为：V
.祝您早日康复                         ---协议字符为：+

以下新增协议为仅报音不显示，显示由用户通过&Cxy+内容＋$，实现，具体见说明书。

'谢谢'                               ---协议字符为：X
'预收'                               ---协议字符为：A
'实收'                               ---协议字符为：B
'找零'                               ---协议字符为：C
'当年帐户余额'                       ---协议字符为：E
'历年帐户余额'                       ---协议字符为：F
'找零请当面点清，谢谢！'             ---协议字符为：G
'帐户余额'                           ---协议字符为：t
'本次消费总额'                       ---协议字符为：u

仅播报金额　　　　　　　　　　　　　 ---协议为：金额＋'p'，如123.32P，播报123.32元

以下为新增到某药房取药的语音协议，仅报音，显示由用户通过&Cxy+内容＋$实现，具体如下：
.请到                                ---协议字符为：<A          
.窗口取药                            ---协议字符为：<B         
.取药                                ---协议字符为：<C           
.西药房                              ---协议字符为：<D            
.中药房                              ---协议字符为：<E          
.1号                                 ---协议字符为：<1              
.2号                                 ---协议字符为：<2      
.3号                                 ---协议字符为：<3          
.4号                                 ---协议字符为：<4           
.5号                                 ---协议字符为：<5         
.6号                                 ---协议字符为：<6         
.7号                                 ---协议字符为：<7         
.8号                                 ---协议字符为：<8          
.9号                                 ---协议字符为：<9          
.10号                                ---协议字符为：<<            
.20号                                ---协议字符为：<>               
如要显示并播报“请到2号窗口取药”，方法为：通过协议&Cxy+内容＋$实现显示，通过报音协议<A、<2、<B连续调用实现发音。   


　　上海通导科技发展有限公司  www.itleader.com.cn
   　　　        2008.1.1
