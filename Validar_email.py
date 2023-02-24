from validate_email import validate_email

from dns.resolver import query
 

def ValidacionDeEmail(Email): 
    validate=Email
    is_valid=validate_email(validate)
    try:
        if(is_valid): 
            domai=validate.rsplit('@',1)[-1]
            final=bool(query(domai,'MX'))
            if final:
                print("Email correcto y verificado")
                return True
        else:
            print("Email invalido")
            return False
    except:
        print("Error de dominio")
        return False


if ValidacionDeEmail("gabriel.dominguez@irtech.com.ar"):
    print("Lo lograste")
else:
    print("No andaaaaaa")