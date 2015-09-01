#include <16f877a.h>
#device *=16 ADC=10 
#include <stdlib.h>
#use delay(clock=20M)
#FUSES NOWDT, HS, NOPUT, NOPROTECT, NODEBUG, NOBROWNOUT, NOLVP, NOCPD, NOWRT
#priority rda,timer1,EXT
#use rs232(baud=9600,parity=n,xmit=pin_c6,rcv=pin_c7)
#byte DAC = 0x08
#byte RF = 0x06
#bit D0 = 0x06.4
#bit D1 = 0x06.5
#bit D2 = 0x06.6
#bit D3 = 0x06.7
//======================================
//Khai bao cac bien
//======================================
int16 xung,i,j,k; //Dem xung Encoder
int8 udk;
char string[10];
char c;
float ki, kp, kpi, kii;//Thong so PI
float Vdat, vantoc, ek, ek2;//Thong so van toc
float udk2, udk3; //Bien trung gian dieu khien
float kd; //Bien trung gian khoi dong
float ampe, ampe2, uik, uik2, volts_ampe; //Thong so dong dien
float temp_V;
int16 temp, adc_value;
//======================================
//Doan chuong trinh ngat dem xung encoder
//======================================
#INT_EXT
void dem_encoder()
{
   xung++;; //Tang Xung len 1
}
//======================================
//Ngat Timer1 - Tinh udk
//======================================
#int_timer1
void ngat_timer1()
{
   //Do toc do
   vantoc = (float)xung*6;
   ek = Vdat - vantoc;
   uik = ampe2 - volts_ampe;
   udk2 = udk2 + kp*ek + ki*ek2 + kpi*uik + kii* uik2;
   ek2 = ek;
   uik2 = uik;
   ampe2 = volts_ampe;
   temp_V = 120/kd;
   if (udk2 > temp_V) udk2 = temp_V;
   if (udk2 < 1) udk2 = 1;
   udk3 = udk2 * kd;
   udk = (int8)udk3;
   DAC = ~udk;
   xung = 0;
   set_timer1(3036);
}
//======================================
//Ngat Truyen thong - nhan ky tu
//======================================
#int_rda
void serial_isr() {
   disable_interrupts(INT_RDA); //Vo hieu ngat truyen thong
   string[i]=getc();
   if (string[i]=='#')
   {
      c=string[i-1];
      i=0;
      switch(c)
            {
         case 'p' :
               {
                  kp=atof(string);break; //Kp toc do
               }
         case 'i' :
               {
                  ki=atof(string);break; //Ki toc do
               }
         case 'q' :
               {
                  kpi=atof(string);break; //Kp dong dien
               }
         case 'j' :
               {
                  kii=atof(string);break; //Ki dong dien
               }
         case 'd' :
               {
                  kd=atof(string);break; //Bien TG khoi dong
               }
         case 'v' : 
               {
                  Vdat = atof(string); break; //Toc do dat
               }
         case '&' :
               {
                  putc('&');break; //Do cong COM
               }
            }
   }
   else i++;
   enable_interrupts(INT_RDA);
}
//======================================
//Chuong trinh con do dong dien
//======================================
void do_dong() 
{
   float temp_volts; //Bien trung gian quy doi ADC sang U
   adc_value = 0;
   k = 100;
   for(j=0;j<k;j++)
   {
      temp = read_adc();
      if (temp >= 512) {
         adc_value += temp;
         delay_ms(2);
      } else {
         k++;
      }
   } 
   volts_ampe = (adc_value/100)-512; //Vi 0A ~ 2,5V ~ adc = 512
   temp_volts = volts_ampe * 0.00488; //Quy doi sang dien ap
   ampe = temp_volts/0.185; //Quy doi sang Ampe
}
#INT_RB
void SyncInt() {
if(D0) putc('#');
if(D1) putc('$');
if(D2) putc('%');
}
//======================================
//Chuong trinh chinh
//======================================
void main()
{
   ext_int_edge(L_TO_H);
   enable_interrupts(int_RB);
   enable_interrupts(int_rda); //Cai dat ngat truyen thong
   ext_int_edge( H_TO_L );
   setup_timer_1(T1_INTERNAL | T1_DIV_BY_8); //Cai dat bo chia Timer 1
   enable_interrupts(INT_TIMER1); //Timer1 65536
   set_timer1(3036); //Cai dat Timer 1 ngat trong 0.1s
   enable_interrupts(INT_EXT);
   enable_interrupts(global);
   //Dat cac gia tri ban dau cho cac bien
   xung = 0;
   vantoc = 0;
   Vdat = 0;
   kp = 0;
   ki = 0;
   kpi = 0;
   kii = 0;
   kd = 0; //Bien TG khoi dong
   //Cai dat vao ra
   set_tris_a(0xFF);
   set_tris_b(0xFF);
   set_tris_c(0xc0);
   set_tris_d(0x00);
   delay_ms(50);
   //Cai dat ADC
   setup_adc(ADC_CLOCK_INTERNAL);
   setup_adc_ports(RA0_ANALOG);
   set_adc_channel(0);
   setup_comparator(NC_NC_NC_NC); 
   setup_vref(FALSE); 
   delay_ms(50);
   while(1)
   {
   //Do dong dien
   do_dong();
   //rf();
   delay_ms(50);
   do_dong();
   //rf();
   delay_ms(50);
   //Xuat len may tinh
   //rf();
   printf("!%4.0f|%1.3f|%u|%4.0f|%4.0f|%u|%u|%u|%u@",(float)vantoc,(float)ampe,udk,(float)ek,(float)Vdat,D0, D1, D2, D3);
  }
}
