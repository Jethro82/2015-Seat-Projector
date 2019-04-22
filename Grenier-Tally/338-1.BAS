OPEN "338-1.txt" FOR INPUT AS #1
CLS

DO
        INPUT #1, a$
        IF LEFT$(a$, 1) = "-" THEN
                PRINT RidingsTotal,
                RidingsTotal = 0
                PRINT MID$(a$, 2, 5);
        ELSE
                RidingsTotal = RidingsTotal + 1
        END IF
        SELECT CASE a$
        CASE "n":
                n = n + 1
        CASE "l":
                l = l + 1
        CASE "c":
                c = c + 1
        CASE "b":
                b = b + 1
        CASE "g":
                g = g + 1
        END SELECT
LOOP UNTIL EOF(1)
PRINT RidingsTotal
PRINT "NDP"; n,
PRINT "CON"; c,
PRINT "LIB"; l,
PRINT "GRN"; g,
PRINT "BlQ"; b
PRINT n + c + l + g + b

