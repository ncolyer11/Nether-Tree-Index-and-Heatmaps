import math
import xlsxwriter

# trust me it was quicker to type this out than screw around with some clean loop (it's not as easy as it looks to
# make a double loop that replicates all this btw)
T = [4, 5, 6, 6, 7, 7, 7, 8, 8, 8, 9, 9, 9, 9, 10, 10, 10, 10, 11, 11, 11, 11, 12, 12, 12, 12, 12, 13, 13, 13, 13, 13,
     14, 14, 14, 14, 14, 16, 16, 16, 16, 16, 16, 18, 18, 18, 18, 18, 18, 18, 20, 20, 20, 20, 20, 20, 20, 22, 22, 22, 22,
     22, 22, 22, 22, 24, 24, 24, 24, 24, 24, 24, 24, 24, 26, 26, 26, 26, 26, 26, 26, 26, 26]
H = [4, 5, 5, 6, 5, 6, 7, 5, 6, 7, 5, 6, 7, 8, 5, 6, 7, 8, 5, 6, 7, 8, 5, 6, 7, 8, 9, 5, 6, 7, 8, 9, 5, 6, 7, 8, 9, 5,
     6, 7, 8, 9, 10, 5, 6, 7, 8, 9, 10, 11, 5, 6, 7, 8, 9, 10, 11, 5, 6, 7, 8, 9, 10, 11, 12, 5, 6, 7, 8, 9, 10, 11, 12,
     13, 5, 6, 7, 8, 9, 10, 11, 12, 13]

M = 0
# trunk and hat height inputs
trunkHeight = 4
hatHeight = 4
outSheet = []
# create file (workbook)
outWorkbook = xlsxwriter.Workbook(r"heatmap.xlsx")

while M < 83:
    trunkHeight = T[M]
    hatHeight = H[M]
    offset = trunkHeight - hatHeight

    # create worksheet
    outSheet.append(outWorkbook.add_worksheet(str(trunkHeight) + "T" + str(hatHeight) + "H"))

    # probability of trunk height (Pt)
    if trunkHeight >= 14:
        Pt = 1 / 120
    elif trunkHeight in [8, 10, 12]:
        Pt = 12 / 120
    else:
        Pt = 11 / 120

    # probability of hat height (Ph)
    if trunkHeight in [4, 5]:
        Ph = 1
    elif trunkHeight == 6 and hatHeight == 5:
        Ph = 1 / 3
    elif trunkHeight == 6 and hatHeight == 6:
        Ph = 2 / 3
    else:
        Ph = 1 / math.floor((trunkHeight + 3) / 3)

    # combined set probability for given trunk and hat height
    P = Pt * Ph

    # coordinate list
    m = 0

    # calculating radius values
    if 8 < hatHeight <= 13:
        air1 = [0] * offset
        constantA = [2, 1, 1]
        twos = [2] * (hatHeight - 6)
        three = [3]
        vinesThree = [3, 3, 3]
        vinesTwo = []
        air2 = [0] * (26 - trunkHeight)
    elif 4 < hatHeight <= 8:
        air1 = [0] * offset
        constantA = [2, 1, 1]
        twos = [2] * (hatHeight - 5)
        three = []
        vinesThree = []
        vinesTwo = [2, 2, 2]
        air2 = [0] * (26 - trunkHeight)
    elif hatHeight == 4:
        air1 = [0] * offset
        constantA = [1, 1]
        twos = []
        three = []
        vinesThree = []
        vinesTwo = [2, 2, 2]
        air2 = [0] * (26 - trunkHeight)
    else:
        raise TypeError("invalid parameter input")

    radii = air1 + vinesTwo + vinesThree + three + twos + constantA + air2

    # looping through every block in the huge fungus volume once & calculating plus storing the probabilities
    Y = 0
    n = 0
    while Y <= 26:
        X = -3
        while X <= 3:
            Z = -3
            while Z <= 3:
                # booleans to simplify region calculations
                bl2 = X == radii[n] or X == -radii[n]  # is on X most edge
                bl3 = Z == radii[n] or Z == -radii[n]  # is on Z most edge
                bl4 = Y <= trunkHeight + 1  # is beneath the max height of the huge fungus + 1
                bl5 = abs(X) <= radii[n]  # is within the max width of the huge fungus
                bl6 = abs(Z) <= radii[n]  # is within the max length of the huge fungus
                bl7 = Y < offset + 3  # is within the bottom 3 layers of the hat (vines region)
                bl8 = Y < trunkHeight  # is beneath the max height of the huge fungus

                if X == 0 and Z == 0 and bl8:
                    L = 100
                    S = 0
                    W = 0
                    region = 'trunk'
                elif (bl2 or bl3) and bl5 and bl6 and bl7:
                    L = 0
                    S = 0
                    W = 27.1125
                    region = 'vines'
                elif bl7 and not (bl2 or bl3) and bl4:
                    L = 0
                    S = 0
                    W = 0
                    region = 'airbottom'
                elif bl2 and bl3 and bl4:
                    L = 0
                    S = 1
                    W = 69.3
                    region = 'corners'
                elif not bl2 and not bl3 and bl4 and bl5 and bl6:
                    L = 0
                    S = 10
                    W = 18
                    region = 'internal'
                elif bl2 != bl3 and bl4 and bl5 and bl6:
                    L = 0
                    S = 0.05
                    W = 97.951
                    region = 'external'
                else:
                    L = 0
                    S = 0
                    W = 0
                    region = 'airtop'

                # not clean sorry but writing data to cells
                outSheet[M].write(m + 1, 0, m)
                outSheet[M].write(m + 1, 1, X)
                outSheet[M].write(m + 1, 2, Y)
                outSheet[M].write(m + 1, 3, Z)
                outSheet[M].write(m + 1, 4, L*P)
                outSheet[M].write(m + 1, 5, S*P)
                outSheet[M].write(m + 1, 6, W*P)
                outSheet[M].write(m + 1, 7, P)
                outSheet[M].write(m + 1, 8, region)

                m += 1
                Z += 1

            X += 1
        Y += 1
        n += 1

    # writing in headers
    outSheet[M].write(0, 0, 'block index')
    outSheet[M].write(0, 1, 'X offset')
    outSheet[M].write(0, 2, 'Y offset')
    outSheet[M].write(0, 3, 'Z offset')
    outSheet[M].write(0, 4, 'stem chance')
    outSheet[M].write(0, 5, 'shroom chance')
    outSheet[M].write(0, 6, 'wart chance')
    outSheet[M].write(0, 7, 'set chance')
    outSheet[M].write(0, 8, 'region')

    print(radii)
    M += 1

outWorkbook.close()
