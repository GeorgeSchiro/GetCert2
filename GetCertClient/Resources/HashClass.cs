using System;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using tvToolbox;

namespace GetCert2
{
    public class HashClass
    {
        public static string sMachineKeyPathFile(tvProfile aoProfile, X509Certificate2 aoCertificate)
        {
            try
            {
                string  lsMachineKeyPathFile = null;
                RSACng  loRSACng = aoCertificate.GetRSAPrivateKey() as RSACng;

                if ( null == loRSACng || null == loRSACng.Key )
                    throw new InvalidOperationException("Unable to obtain CNG private key (for key file cleanup).");

                if ( !String.IsNullOrEmpty(loRSACng.Key.UniqueName) )
                {
                    lsMachineKeyPathFile = Path.Combine(@"C:\ProgramData\Microsoft\Crypto\Keys", loRSACng.Key.UniqueName);

                    if ( !File.Exists(lsMachineKeyPathFile) )
                        throw new InvalidOperationException(String.Format("The purported CNG private key file (\"{0}\" - for cleanup) does not exist!", lsMachineKeyPathFile));
                }

                return lsMachineKeyPathFile;
            }
            catch
            {
                if ( aoProfile.bValue("-DeclareErrorOnPrivateKeyFileCleanup", false) )
                    throw;

                return null;
            }
        }

        public static string sEncryptedBase64(X509Certificate2 aoCertificate, string asClearText)
        {
            if ( null == aoCertificate )
                throw new InvalidOperationException("The given certificate does not exist!");

            RSA loRSA = aoCertificate.GetRSAPublicKey();

            return Convert.ToBase64String(loRSA.Encrypt(ASCIIEncoding.UTF8.GetBytes(asClearText), RSAEncryptionPadding.OaepSHA512));
        }

        public static string sDecrypted(X509Certificate2 aoCertificate, string asEncryptedBase64)
        {
            if ( null == aoCertificate )
                throw new InvalidOperationException("The given certificate does not exist!");
            
            RSA loRSA = aoCertificate.GetRSAPrivateKey();

            if ( null == loRSA )
                throw new InvalidOperationException("The given certificate has no private key!");

            return ASCIIEncoding.UTF8.GetString(loRSA.Decrypt(Convert.FromBase64String(asEncryptedBase64), RSAEncryptionPadding.OaepSHA512));
        }

        public static string sHashIt(tvProfile aoProfile)
        {
            return HashClass.sHashPw(aoProfile);
        }

        public static string sHashPw(tvProfile aoProfile)
        {
            StringBuilder   lsbHashPw = new StringBuilder();
            byte[]          lbtArray = SHA256.Create().ComputeHash(Encoding.UTF8.GetBytes(aoProfile.ToString()));
                            foreach (byte lbtValue in lbtArray)
                                lsbHashPw.Append(lbtValue.ToString());

            return lsbHashPw.ToString();
        }
    }
}
