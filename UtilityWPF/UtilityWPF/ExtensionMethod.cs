using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace UtilityWPF
{
    static class ExtensionMethod
    {
        //public static RSA GetRSAPrivateKeyWithPin(this string pin)
        //{
        //    RSA rsa = cert.GetRSAPrivateKey();

        //    //if (rsa is RSACryptoServiceProvider rsaCsp)
        //    //{
        //    //    // Current code
        //    //    SetPin(rsaCsp);
        //    //    return rsa;
        //    //}

        //    if (rsa is RSACng rsaCng)
        //    {
        //        // Set the PIN, an explicit null terminator is required to this Unicode/UCS-2 string.

        //        byte[] propertyBytes;

        //        if (pin[pin.Length - 1] == '\0')
        //        {
        //            propertyBytes = Encoding.Unicode.GetBytes(pin);
        //        }
        //        else
        //        {
        //            propertyBytes = new byte[Encoding.Unicode.GetByteCount(pin) + 2];
        //            Encoding.Unicode.GetBytes(pin, 0, pin.Length, propertyBytes, 0);
        //        }

        //        const string NCRYPT_PIN_PROPERTY = "965539";

        //        CngProperty pinProperty = new CngProperty(
        //            NCRYPT_PIN_PROPERTY,
        //            propertyBytes,
        //            CngPropertyOptions.None);

        //        rsaCng.Key.SetProperty(pinProperty);
        //        return rsa;
        //    }

        //    // If you're on macOS or Linux neither of the above will hit. There's
        //    // also no standard model for setting a PIN on either of those OS families.
        //    rsa.Dispose();
        //    throw new NotSupportedException($"Don't know how to set the PIN for {rsa.GetType().FullName}");
        //}
    }
}
