library(readxl)

data<-as.data.frame(read_excel("BookingDEC29.xlsx",sheet = "SHEET2"))
data2 <- as.data.frame(read_excel("BookingDEC29.xlsx",sheet = "Sheet"))

str(data)
str(data2)
library(ggplot2)

data$總分 <- as.numeric(as.character(data$總分))
data$員工素質 <- as.numeric(as.character(data$員工素質))
data$設施 <- as.numeric(as.character(data$設施))
data$清潔程度 <- as.numeric(as.character(data$清潔程度))
data$舒適程度 <- as.numeric(as.character(data$舒適程度))
data$性價比 <- as.numeric(as.character(data$性價比))
data$住宿地點 <- as.numeric(as.character(data$住宿地點))
data$`免費 WiFi` <- as.numeric(data$`免費 WiFi`)
data2$人數 <-as.numeric(as.character(data2$人數))
data2$原始價格 <-as.numeric(as.character(data2$原始價格))
data2$目前價格 <-as.numeric(as.character(data2$目前價格))

str(data)
str(data2)

data$星級[which(data$星級 == 8)] = 0
data$星級[which(data$星級 == 6)] = 0
#
data <- data[!duplicated(data$旅館名稱), ]
#289
data <-  subset(data,data$旅館名稱 !="AJ Residence 安捷國際公寓酒店")


pl <-ggplot(data,aes(x=`總分`))
plot1<-pl+geom_histogram(binwidth = 0.1, color="#1E3815",fill="#69b3a2")+labs(x = '總分',y="數量")+theme_bw()
plot1
pl2 <-ggplot(data,aes(x=`員工素質`))
plot2<-pl2+geom_histogram(binwidth = 0.1, color="#1E3815",fill="#69b3a2")+labs(y="數量")+theme_bw()

pl3 <-ggplot(data,aes(x=`設施`))
plot3<-pl3+geom_histogram(binwidth = 0.1, color="#1E3815",fill="#69b3a2")+labs(y="數量")+theme_bw()

pl4 <-ggplot(data,aes(x=`清潔程度`))
plot4<-pl4+geom_histogram(binwidth = 0.1, color="#1E3815",fill="#69b3a2")+labs(y="數量")+theme_bw()

pl5 <-ggplot(data,aes(x=`舒適程度`))
plot5<-pl5+geom_histogram(binwidth = 0.1, color="#1E3815",fill="#69b3a2")+labs(y="數量")+theme_bw()

pl6 <-ggplot(data,aes(x=`性價比`))
plot6<-pl6+geom_histogram(binwidth = 0.1, color="#1E3815",fill="#69b3a2")+labs(y="數量")+theme_bw()

pl7 <-ggplot(data,aes(x=`住宿地點`))
plot7<-pl7+geom_histogram(binwidth = 0.1, color="#1E3815",fill="#69b3a2")+labs(y="數量")+theme_bw()

pl8 <-ggplot(data,aes(x=`免費 WiFi`))
plot8<-pl8+geom_histogram(binwidth = 0.1, color="#1E3815",fill="#69b3a2")+labs(y="數量")+theme_bw()

library(patchwork)
plot1+plot2+plot3+plot4+plot5+plot6+plot7+plot8

##########################總分-評分
#279
data <-subset(data,data$總分>=7)


pla<-ggplot(data,aes(x =`員工素質` ,y=`總分`))
plot9<-pla + geom_point(color='DarkCyan')+geom_smooth(method = "lm",color='brown',formula= y~x)
plot9

plb<-ggplot(data,aes(x =`設施` ,y=`總分`))
plot10<-plb + geom_point(color='DarkCyan')+geom_smooth(method = "lm",color='brown',formula= y~x)
plot10

plc<-ggplot(data,aes(x =`清潔程度` ,y=`總分`))
plot11<-plc + geom_point(color='DarkCyan')+geom_smooth(method = "lm",color='brown',formula= y~x)

pld<-ggplot(data,aes(x =`舒適程度` ,y=`總分`))
plot12<-pld + geom_point(color='DarkCyan')+geom_smooth(method = "lm",color='brown',formula= y~x)

ple<-ggplot(data,aes(x =`性價比`,y=`總分`))
plot13<-ple + geom_point(color='DarkCyan')+geom_smooth(method = "lm",color='brown',formula= y~x)

ple<-ggplot(data,aes(x =`住宿地點` ,y=`總分`))
plot14<-ple + geom_point(color='DarkCyan')+geom_smooth(method = "lm",color='brown',formula= y~x)

ple<-ggplot(data,aes(x =`免費 WiFi` ,y=`總分`))
plot15<-ple + geom_point(color='DarkCyan')+geom_smooth(method = "lm",color='brown',formula= y~x)

plot9+plot10+plot11+plot12+plot13+plot14+plot15

########################## 星級-評分
data$星級 <- as.factor(data$星級)
#remove one star hotel(only one data)
data<-subset(data,data$星級!=1)
#significant level
library(ggsignif)

bx1 <- ggplot(data,aes(x=`星級`,y=`總分`,color=`星級`)) #bx1 + geom_point()
bxplot1<-bx1 + geom_boxplot()+
  stat_summary(fun.y=mean, geom="point", shape=16, size=1, col='DarkRed')+
  geom_signif(comparisons = list(c("0", "2"),c("2", "3"),c("3", "4"),c("4", "5")),
              y_position = c(9.5,9.7,9.9,10.1),col = 1,
              map_signif_level=TRUE)+
  ylim(bottom =min(data$總分) ,top=10.5)+guides(color=F)
bxplot1

bx2 <- ggplot(data,aes(x=`星級`,y=`員工素質`,color=`星級`))
bxplot2<-bx2 + geom_boxplot()+
  stat_summary(fun.y=mean, geom="point", shape=16, size=1, col='DarkRed')+
  geom_signif(comparisons = list(c("0", "2"),c("2", "3"),c("3", "4"),c("4", "5")),
              y_position = c(9.5,9.7,9.9,10.1),col = 1,
              map_signif_level=TRUE)+
  ylim(bottom =min(data$員工素質) ,top=10.5)+guides(color=F)
  

bx3 <- ggplot(data,aes(x=`星級`,y=`設施`,color=`星級`))
bxplot3<-bx3 + geom_boxplot()+
  stat_summary(fun.y=mean, geom="point", shape=16, size=1, col='DarkRed')+
  geom_signif(comparisons = list(c("0", "2"),c("2", "3"),c("3", "4"),c("4", "5")),
              y_position = c(9.5,9.7,9.9,10.1),col = 1,
              map_signif_level=TRUE)+
  ylim(bottom =min(data$設施) ,top=10.5)+guides(color=F)

bx4 <- ggplot(data,aes(x=`星級`,y=`清潔程度`,color=`星級`))
bxplot4<-bx4 + geom_boxplot()+
  stat_summary(fun.y=mean, geom="point", shape=16, size=1, col='DarkRed')+
  geom_signif(comparisons = list(c("0", "2"),c("2", "3"),c("3", "4"),c("4", "5")),
              y_position = c(9.5,9.7,9.9,10.1),col = 1,
              map_signif_level=TRUE)+
  ylim(bottom =min(data$清潔程度) ,top=10.5)+guides(color=F)

bx5 <- ggplot(data,aes(x=`星級`,y=`舒適程度`,color=`星級`))
bxplot5<-bx5 + geom_boxplot()+
  stat_summary(fun.y=mean, geom="point", shape=16, size=1, col='DarkRed')+
  geom_signif(comparisons = list(c("0", "2"),c("2", "3"),c("3", "4"),c("4", "5")),
              y_position = c(9.5,9.7,9.9,10.1),col = 1,
              map_signif_level=TRUE)+
  ylim(bottom =min(data$舒適程度) ,top=10.5)+guides(color=F)

bx6 <- ggplot(data,aes(x=`星級`,y=`性價比`,color=`星級`))
bxplot6<-bx6 + geom_boxplot()+
  stat_summary(fun.y=mean, geom="point", shape=16, size=1, col='DarkRed')+
  geom_signif(comparisons = list(c("0", "2"),c("2", "3"),c("3", "4"),c("4", "5")),
              y_position = c(9.5,9.7,9.9,10.1),col = 1,
              map_signif_level=TRUE)+
  ylim(bottom =min(data$性價比) ,top=10.5)+guides(color=F)

bx7 <- ggplot(data,aes(x=`星級`,y=`住宿地點`,color=`星級`))
bxplot7<-bx7 + geom_boxplot()+
  stat_summary(fun.y=mean, geom="point", shape=16, size=1, col='DarkRed')+
  geom_signif(comparisons = list(c("0", "2"),c("2", "3"),c("3", "4"),c("4", "5")),
              y_position = c(9.5,9.7,9.9,10.1),col = 1,
              map_signif_level=TRUE)+
  ylim(bottom =min(data$住宿地點) ,top=10.5)+guides(color=F)

bx8 <- ggplot(data,aes(x=`星級`,y=`免費 WiFi`,color=`星級`))
bxplot8<-bx8 + geom_boxplot()+
  stat_summary(fun.y=mean, geom="point", shape=16, size=1, col='DarkRed')+
  geom_signif(comparisons = list(c("0", "2"),c("2", "3"),c("3", "4"),c("4", "5")),
              y_position = c(9.5,9.7,9.9,10.1),col = 1,
              map_signif_level=TRUE)+
  ylim(bottom =min(data$`免費 WiFi`) ,top=10.5)+guides(color=F)


bxplot1+bxplot2+bxplot3+bxplot4+bxplot5+bxplot6+bxplot7+bxplot8



##########################組間差異
dat.aov <- aov(總分 ~ 星級, data=df_all)
summary(dat.aov)

dat.aov <- aov(員工素質 ~ 星級, data=df_all)
summary(dat.aov)

dat.aov <- aov(設施 ~ 星級, data=df_all)
summary(dat.aov)

dat.aov <- aov(清潔程度 ~ 星級, data=df_all)
summary(dat.aov)

dat.aov <- aov(舒適程度 ~ 星級, data=df_all)
summary(dat.aov)

dat.aov <- aov(性價比 ~ 星級, data=df_all)
summary(dat.aov)

dat.aov <- aov(住宿地點 ~ 星級, data=df_all)
summary(dat.aov)

dat.aov <- aov(data$`免費 WiFi` ~ 星級, data=df_all)
summary(dat.aov)



##########################
#整理各飯店兩人房最低價格
#df_nna %>% group_by(month) %>% summarise(sum_of_every_month = sum(distance,na.rm = T))
library(dplyr)
data2$目前價格<-as.integer(data2$目前價格)
data2.1 <- filter(data2,人數==2)
datahotel <-summarize(group_by(data2.1,旅館名稱),'價格'=min(目前價格))
datahotel$價格 <-as.numeric(datahotel$價格)



#Merge datahotel and data (right join)
df_all <-merge(x = datahotel, y = data, by = "旅館名稱",all.y = T)

#sum(is.na(df_all$a)) #總共有幾間是na(沒有兩人房)

#刪除價格過高的
df_all <-filter(df_all,df_all$價格<10000)

pl <- ggplot(df_all,aes(x =`總分` ,y =`價格`))
P1<-pl+geom_point(color='DarkCyan')+geom_smooth(method = lm,color='brown',formula= y~x)
P1

pl <- ggplot(df_all,aes(x =`設施` ,y =`價格`))
P2<-pl+geom_point(color='DarkCyan')+geom_smooth(method = lm,color='brown',formula= y~x)
#分數低時:提升設施分數無明顯影響 分數高時:提升設施可以提高價格


pl <- ggplot(df_all,aes(x =`清潔程度` ,y =`價格`))
P3<-pl+geom_point(color='DarkCyan')+geom_smooth(method = lm,color='brown',formula= y~x)

pl <- ggplot(df_all,aes(x =`員工素質` ,y =`價格`))
P4<-pl+geom_point(color='DarkCyan')+geom_smooth(method = lm,color='brown',formula= y~x)

pl <- ggplot(df_all,aes(x =`舒適程度` ,y =`價格`))
P5<-pl+geom_point(color='DarkCyan')+geom_smooth(method = lm,color='brown',formula= y~x)

pl <- ggplot(df_all,aes(x =`價格` ,y =`性價比`))
P6<-pl+geom_point(color='DarkCyan')+geom_smooth(method = lm,color='brown',formula= y~x)

pl <- ggplot(df_all,aes(x =`住宿地點` ,y =`價格`))
P7<-pl+geom_point(color='DarkCyan')+geom_smooth(method = lm,color='brown',formula= y~x)

pl <- ggplot(df_all,aes(x =`免費 WiFi` ,y =`價格`))
P8<-pl+geom_point(color='DarkCyan')+geom_smooth(method = lm,color='brown',formula= y~x)

P1+P2+P3+P4+P5+P6+P7+P8

##########################

pl <- ggplot(df_all,aes(x =`總分` ,y =`價格`))
pl+geom_point(color='DarkSlateGray ')+
  geom_smooth(method = lm,color='Coral',formula= y~x)+
  theme_bw()+xlim(bottom=min(df_all$總分),top=max(df_all$總分))





compare_means(len ~ dose,  data = ToothGrowth, ref.group = "0.5",
              method = "t.test")


shapiro.test(data$`免費 WiFi`)

#########################
#Adjust prrice to norm
df_all$newprice = log((df_all$價格)+1)
#兩兩相關性
library(corrr)
corr_df_all<-correlate(df_all)

library(GGally)
GGally::ggpairs(data = df_all[3:10])

df_all[1]
#Caculate the varibles that influence price the most(drop those in high corr)
mlm <- lm(newprice ~ 舒適程度+性價比+住宿地點, data = df_all)
summary(mlm)

########Conclusion:
# Given 舒適程度x1、性價比x2、住宿地點x3
# With about 60% confidence to predict the price by the formula:
# Price_log =  (x1 *0.77527 + x2 * (-0.68024) + x3 * 0.26457)+4.85157
# Price_prediction = expm1(Price_log)

