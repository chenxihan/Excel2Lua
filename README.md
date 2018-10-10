# Excel2Lua
表格前4行为固定格式：
第一行：备注单列的含义，不会导到lua中
第二行：类型（int,double,string,tabint,tabdouble,tabstring）
第三行：键名（必须有一个为k，建议第一列为k）
第四行：s 服务器用；c 客户端用；m 转成表，用|分割；e 忽略空单元（不填e则不能为空单元）;--s、c、m、e安需求组合，eg:sc(客户端和服务器都要用)，sce(客户端和服务器都要用，且忽略空单元)