/*
Purpose:  Converte uma string de caracteres em uma string 128B ou um 128C  de c¢digo de barras.
Input  :  string ou variavel contendo uma string para convers∆o
Output :  string de c¢digo de barras 
Syntax:  Cb128(c-string)
basic usage:
//seta a fonte Code 128 no range de cÇlulas
    chWorkSheet:Range("K1:K18"):FONT:NAME = "Code 128".
//chama a funá∆o que far† a convers∆o da string
    chWorksheet:range("K1:K18"):VALUE = Cb128(c-string).
*/

FUNCTION Cb128TestSum RETURNS INTEGER (pi_i AS INT, pi_min AS INT, pi_chaine AS CHAR):
    /*se os caracteres pi_min de pi_i forem numÇricos, ent∆o pi_min = 0*/
    ASSIGN pi_min = pi_min - 1.
    IF pi_i + pi_min <= LENGTH(pi_chaine) THEN DO WHILE pi_min >= 0:
        IF ASC(SUBSTRING(pi_chaine, pi_i + pi_min, 1)) < 48 OR
           ASC(SUBSTRING(pi_chaine, pi_i + pi_min, 1)) > 57 THEN DO:
            RETURN pi_min.
        END.
        ELSE DO:
            ASSIGN pi_min = pi_min - 1.
        END.
    END.
    RETURN pi_min.
END.

/* Fonction π appeler */
FUNCTION Cb128 RETURNS CHAR (pi_chaine AS CHAR):

  DEF VAR vi_i        AS INT  NO-UNDO.
  DEF VAR vi_min      AS INT  NO-UNDO.
  DEF VAR vi_dummy    AS INT  NO-UNDO.
  DEF VAR vi_checksum AS INT  NO-UNDO.
  DEF VAR vl_TableB   AS LOG  NO-UNDO INIT TRUE.
  DEF VAR vc_Cb128    AS CHAR NO-UNDO.

  IF LENGTH(pi_chaine) > 0 THEN DO:
  
      /* Verifica se os caracteres s∆o v†lidos */
      DO vi_i = 1 TO LENGTH(pi_chaine):
          IF ASC(SUBSTRING(pi_chaine, vi_i, 1)) < 32 OR
             ASC(SUBSTRING(pi_chaine, vi_i, 1)) > 126 THEN DO:
                RETURN "".
          END.
      END.

        
      /* C†lculo da string de c¢digo com uso otimizado das tabelas B e C */
      ASSIGN vi_i = 1.
      DO WHILE vi_i <= LENGTH(pi_chaine):
          /* Veja se Ç interessante mudar para a tabela C
             sim para 4 d°gitos no in°cio ou no final, sen∆o se 6 d°gitos */          
          IF vl_TableB THEN DO:
             ASSIGN vi_min = IF (vi_i = 1 OR vi_i + 3 = LENGTH(pi_chaine)) THEN 4 ELSE 6.
             ASSIGN vi_min = Cb128TestSum(vi_i, vi_min, pi_chaine).
             IF vi_min < 0 THEN DO: /*Escolha da tabela C*/ 
                IF vi_i = 1 THEN DO: /*iniciando com a tabela C*/
                    ASSIGN vc_Cb128 = CHR(205).
                END.
                ELSE DO: /* troca para tabela C */
                    ASSIGN vc_Cb128 = vc_Cb128 + CHR(199).
                END.
                ASSIGN vl_TableB = False.
             END.
             ELSE DO:
                 IF vi_i = 1 THEN DO: /* iniciando coma tabela B */
                     ASSIGN vc_Cb128 = CHR(204).
                 END.
             END.
          END.

          /* Processamento com 2 digitos na tabela C */
          IF NOT vl_TableB THEN DO:
              ASSIGN vi_min = 2.
              ASSIGN vi_min = Cb128TestSum(vi_i, vi_min, pi_chaine).
              IF vi_min < 0 THEN DO: /*Ok para 2 digitos, continuar processamento */
                  ASSIGN vi_dummy = INTEGER(SUBSTRING(pi_chaine, vi_i, 2)).
                  ASSIGN vi_dummy = IF vi_dummy < 95 THEN vi_dummy + 32 ELSE vi_dummy + 100.
                  ASSIGN vc_Cb128 = vc_Cb128 + CHR(vi_dummy)
                         vi_i = vi_i + 2.
              END.
              ELSE DO: /* sem 2 digitos altera para tabela B */
                  ASSIGN vc_Cb128  = vc_Cb128 + CHR(200)
                         vl_TableB = TRUE.
              END.
          END.

          /* Processa 1 digito na tabela B */
          IF vl_TableB THEN DO: 
            ASSIGN vc_Cb128 = vc_Cb128 + SUBSTRING(pi_chaine, vi_i, 1)
                   vi_i     = vi_i + 1.
          END.
      END. /* DO WHILE  */


      /* calculo de controle(checksum) */
      DO vi_i = 1 TO LENGTH(vc_Cb128):
          ASSIGN vi_dummy = ASC(SUBSTRING(vc_Cb128, vi_i, 1)).
          ASSIGN vi_dummy = IF vi_dummy < 127 THEN vi_dummy - 32 ELSE vi_dummy - 100.
          IF vi_i = 1 THEN DO:
              ASSIGN vi_Checksum = vi_dummy.
          END.
          ASSIGN vi_checksum = (vi_checksum + (vi_i - 1) * vi_dummy) MODULO 103.
      END.
    
    /* calculo do checksum da tabela ASCII */
    ASSIGN vi_CheckSum = IF vi_CheckSum < 95 THEN vi_CheckSum + 32 ELSE vi_CheckSum + 100.


    /* Adicionar checksum e finaliza */
    ASSIGN vc_Cb128 = vc_Cb128 + CHR(vi_CheckSum) + CHR(206).
  END. /* IF */

  RETURN vc_Cb128.

END FUNCTION.
