Code Explanation (English)
This code is designed to generates discount curves for given OIS (Overnight Index Swap) and CDS (Credit Default Swap) data. It focuses on two main components: calculating risk-free interest rates from OIS data and calculating liquidity premiums from CDS data. 

Main tool used: scipy.optimize.minimize

Data Reading and Filtering:
First, the OIS and CDS data are read from CSV files. OIS data is filtered by specific currencies (EUR, USD, TRY, GBP), while CDS data is filtered by specific restructuring types.

Data Transformation:
The maturity information (tenor) can be in weeks, months, or years. Hence, the maturity data is converted to years. The maturity and yield information are used to calculate risk-free interest rates and liquidity premiums.

Nelson-Siegel-Svensson (NSS) Model:
The Nelson-Siegel-Svensson (NSS) model is used to calculate risk-free interest rates. This model describes the yield curve using a four-term function.
The parameters of the NSS model are optimized using the scipy.optimize.minimize function. Various initial guesses and optimization methods are tried to find the parameters that yield the lowest error.

Liquidity Premium Calculation:
Liquidity premiums are calculated using CDS data. This process is similar to the optimization performed for risk-free interest rates using OIS data.
The liquidity premium is treated as a risk factor added to the risk-free interest rates. It is also optimized using the NSS model.

Saving the Results:
The calculated risk-free rates and liquidity premiums are calculated for specific days and maturities. These results are then saved to an Oracle database and Excel files.
Separate Excel files are generated for each currency, and the results are written to different sheets.

Email Notification:
After the results are calculated, all files in the specified folder are collected and sent to a recipient via an Outlook email.
This code combines the use of advanced financial modeling techniques with practical data processing and storage, ensuring that the discount curves are optimized and communicated efficiently.



Kod Açıklaması (Türkçe) 
Bu kod, verilen OIS (Overnight Index Swap) ve CDS (Credit Default Swap) verileri için iskonto eğrilerini optimize etmek amacıyla tasarlanmıştır. Kod iki ana bileşen içerir: OIS verileri ile risksiz faiz oranlarını hesaplama ve CDS verileri ile likidite primini hesaplama.

Veri Okuma ve Filtreleme:
İlk olarak, OIS ve CDS verileri CSV dosyalarından okunur. OIS verileri belirli para birimlerine göre filtrelenir (EUR, USD, TRY, GBP). CDS verileri ise belirli yeniden yapılandırma türlerine göre filtrelenir.

Veri Dönüştürme:
Vade (maturity) bilgileri haftalar, aylar veya yıllar cinsinden olabilir. Bu nedenle vade bilgileri yıllık süreye dönüştürülür. Vade ve getiri bilgileri risksiz faiz oranlarını ve likidite primlerini hesaplamak için kullanılır.

Nelson-Siegel-Svensson (NSS) Modeli:
Risksiz faiz oranlarını hesaplamak için Nelson-Siegel-Svensson (NSS) modeli kullanılır. Bu model, faiz eğrisini dört terimli bir fonksiyonla açıklar.
NSS modelindeki parametrelerin optimize edilmesi için scipy.optimize.minimize fonksiyonu kullanılır. Çeşitli başlangıç tahminleri ve optimizasyon yöntemleri denenerek, en düşük hatayı veren parametreler bulunur.

Likidite Primi Hesaplama:
CDS verileri kullanılarak likidite primi hesaplanır. Bu süreç, OIS verileriyle risksiz faiz oranları için yapılan optimizasyona benzer şekilde gerçekleştirilir.
Likidite primi, risksiz faiz oranlarına eklenen bir risk faktörü olarak ele alınır. Yine NSS modeli kullanılarak optimize edilir.

Sonuçların Kaydedilmesi:
Hesaplanan risksiz faiz oranları ve likidite primi, belirli günler ve vadeler için hesaplanır. Bu veriler daha sonra bir Oracle veritabanına ve Excel dosyalarına kaydedilir.
Her para birimi için ayrı ayrı Excel dosyaları oluşturulur ve sonuçlar farklı sayfalara yazılır.

E-posta Gönderme:
Sonuçlar hesaplandıktan sonra, belirlenen klasördeki dosyalar toplanır ve bir Outlook e-postası aracılığıyla belirtilen alıcıya gönderilir.
