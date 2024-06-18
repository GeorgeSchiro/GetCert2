using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using tvToolbox;

namespace GetCert2
{
    public class HashClass
    {
        // From https://gist.github.com/haf/5629584

        [DllImport("crypt32", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern bool CryptAcquireCertificatePrivateKey(IntPtr pCert, uint dwFlags, IntPtr pvReserved,
                                                                    ref IntPtr phCryptProv, ref int pdwKeySpec,
                                                                    ref bool pfCallerFreeProv);

        [DllImport("advapi32", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern bool CryptGetProvParam(IntPtr hCryptProv, CryptGetProvParamType dwParam, IntPtr pvData,
                                                    ref int pcbData, uint dwFlags);

        [DllImport("advapi32", SetLastError = true)]
        internal static extern bool CryptReleaseContext(IntPtr hProv, uint dwFlags);

        internal enum CryptGetProvParamType
        {
            PP_ENUMALGS = 1,
            PP_ENUMCONTAINERS = 2,
            PP_IMPTYPE = 3,
            PP_NAME = 4,
            PP_VERSION = 5,
            PP_CONTAINER = 6,
            PP_CHANGE_PASSWORD = 7,
            PP_KEYSET_SEC_DESCR = 8, // get/set security descriptor of keyset
            PP_CERTCHAIN = 9, // for retrieving certificates from tokens
            PP_KEY_TYPE_SUBTYPE = 10,
            PP_PROVTYPE = 16,
            PP_KEYSTORAGE = 17,
            PP_APPLI_CERT = 18,
            PP_SYM_KEYSIZE = 19,
            PP_SESSION_KEYSIZE = 20,
            PP_UI_PROMPT = 21,
            PP_ENUMALGS_EX = 22,
            PP_ENUMMANDROOTS = 25,
            PP_ENUMELECTROOTS = 26,
            PP_KEYSET_TYPE = 27,
            PP_ADMIN_PIN = 31,
            PP_KEYEXCHANGE_PIN = 32,
            PP_SIGNATURE_PIN = 33,
            PP_SIG_KEYSIZE_INC = 34,
            PP_KEYX_KEYSIZE_INC = 35,
            PP_UNIQUE_CONTAINER = 36,
            PP_SGC_INFO = 37,
            PP_USE_HARDWARE_RNG = 38,
            PP_KEYSPEC = 39,
            PP_ENUMEX_SIGNING_PROT = 40,
            PP_CRYPT_COUNT_KEY_USE = 41,
        }

        public static string sMachineKeyPathFile(tvProfile aoProfile, X509Certificate2 aoCertificate)
        {
            string lsMachineKeyPathFile = null;

            try
            {
                lsMachineKeyPathFile = Path.Combine(aoProfile.sValue("-MachineKeysPath", @"C:\ProgramData\Microsoft\Crypto\RSA\MachineKeys")
                                        , HashClass.GetKeyFileName(aoCertificate));
            }
            catch (InvalidOperationException ex)
            {
                if ( aoProfile.bValue("-DeclareErrorOnPrivateKeyFileCleanup", false) )
                    throw ex;
            }

            return lsMachineKeyPathFile;
        }

        private static string GetKeyFileName(X509Certificate2 cert)
        {
            var hProvider = IntPtr.Zero; // CSP handle
            var freeProvider = false; // Do we need to free the CSP ?
            uint acquireFlags = 0;
            var _keyNumber = 0;
            string keyFileName = null;
            byte[] keyFileBytes = null;

            //
            // Determine whether there is private key information available for this certificate in the key store
            //
            if (CryptAcquireCertificatePrivateKey(cert.Handle,
                                                acquireFlags,
                                                IntPtr.Zero,
                                                ref hProvider,
                                                ref _keyNumber,
                                                ref freeProvider))
            {
                var pBytes = IntPtr.Zero; // Native Memory for the CRYPT_KEY_PROV_INFO structure
                var cbBytes = 0; // Native Memory size

                try
                {
                    if (CryptGetProvParam(hProvider, CryptGetProvParamType.PP_UNIQUE_CONTAINER, IntPtr.Zero, ref cbBytes, 0))
                    {
                        pBytes = Marshal.AllocHGlobal(cbBytes);

                        if (CryptGetProvParam(hProvider, CryptGetProvParamType.PP_UNIQUE_CONTAINER, pBytes, ref cbBytes, 0))
                        {
                            keyFileBytes = new byte[cbBytes];

                            Marshal.Copy(pBytes, keyFileBytes, 0, cbBytes);

                            // Copy eveything except tailing null byte
                            keyFileName = Encoding.ASCII.GetString(keyFileBytes, 0, keyFileBytes.Length - 1);
                        }
                    }
                }
                finally
                {
                    if (freeProvider)
                        CryptReleaseContext(hProvider, 0);

                    //
                    // Free our native memory
                    //
                    if (pBytes != IntPtr.Zero)
                        Marshal.FreeHGlobal(pBytes);
                }
            }

            if (keyFileName == null)
                throw new InvalidOperationException("Unable to obtain private key file name");

            return keyFileName;
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
