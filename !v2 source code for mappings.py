import math
import xlsxwriter

bb = 0
Block = 0
T = list(range(4, 14)) + list(range(14, 27, 2))
outSheet = []
# create file (workbook)
outWorkbook = xlsxwriter.Workbook(r"HeatMapV2.xlsx")

for count1, trunkHeight in enumerate(T):
    if trunkHeight in [4, 5]:
        base = trunkHeight
        offset1 = 1
        H = list(range(base, base + offset1))
    elif trunkHeight == 6:
        H = [5, 6, 6]
    else:
        base = 5
        offset1 = math.floor(1 + trunkHeight / 3)
        H = list(range(base, base + offset1))

    for hatHeight in H:
        # probability of trunk height (Pt)
        Pt = (
            1 / 120 if trunkHeight > 13 else
            12 / 120 if trunkHeight in [8, 10, 12] else
            11 / 120
        )
        # probability of hat height (Ph)
        Ph = (
            1 if trunkHeight in [4, 5] else
            1 / math.floor(1 + trunkHeight / 3)
        )

        offset2 = trunkHeight - hatHeight
        LVS = []
        Fringe = [[1, 1], [1, 2], [2, 1], [2, 2]]
        PFringe = [2/9, 1/9, 4/9, 2/9]
        NFringe = ["A", "B", "C", "D"]

        N = 0
        while N < 4:
            m = 0
            if hatHeight == 4:
                LVS = [2, 2] + Fringe[N] + [1] + [0]*(26 - trunkHeight)
                Pf = PFringe[N]
            elif 4 < hatHeight <= 8:
                LVS = [0]*offset2 + [2]*(hatHeight - 2) + Fringe[N] + [1] + [0]*(26 - trunkHeight)
                Pf = PFringe[N]
            elif 8 < hatHeight <= 13:
                LVS = [0]*offset2 + [3]*4 + [2]*(hatHeight - 6) + Fringe[N] + [1] + [0]*(26 - trunkHeight)
                Pf = PFringe[N]
            else:
                raise TypeError("invalid parameter input")

            P = Pt * Ph * Pf

            Y = 0
            # create worksheet
            outSheet.append(outWorkbook.add_worksheet(f"T{trunkHeight}H{hatHeight}({bb}{NFringe[N]})"))
            while Y <= 26:
                X = -3
                while X <= 3:
                    Z = -3
                    while Z <= 3:
                        # booleans to simplify region calculations
                        bl2 = abs(X) == LVS[Y]  # is on Z most edge
                        bl3 = abs(Z) == LVS[Y]  # is on Z most edge
                        bl4 = Y <= trunkHeight + 1  # is beneath the max height of the huge fungus + 1
                        bl5 = abs(X) <= LVS[Y]  # is within the max width of the huge fungus
                        bl6 = abs(Z) <= LVS[Y]  # is within the max length of the huge fungus
                        bl7 = Y == offset2  # first vines region layer
                        bl8 = Y == offset2 + 1  # second vines region layer
                        bl9 = Y == offset2 + 2  # third vines region layer
                        bl10 = Y < trunkHeight  # is beneath the max height of the huge fungus
                        bl11 = Y <= trunkHeight  # is beneath or equal to the max height of the huge fungus

                        if X == 0 and Z == 0 and bl10:
                            L = 1
                            S = 0
                            W = 0
                            region = 'trunk'
                        elif (bl2 or bl3) and bl5 and bl6 and bl7:
                            L = 0
                            S = 0
                            W = 0.15
                            region = 'vines1'
                        elif (bl2 or bl3) and bl5 and bl6 and bl8:
                            L = 0
                            S = 0
                            W = 0.2775
                            region = 'vines2'
                        elif (bl2 or bl3) and bl5 and bl6 and bl9:
                            L = 0
                            S = 0
                            W = 0.385875
                            region = 'vines3'
                        elif bl7 and not (bl2 or bl3) and bl4:
                            L = 0
                            S = 0
                            W = 0
                            region = 'air1'
                        elif bl2 and bl3 and bl11:
                            L = 0
                            S = 0.01
                            W = 0.693
                            region = 'corners'
                        elif not bl2 and not bl3 and bl5 and bl6 and not bl7 and not bl8 and not bl9 and bl10:
                            L = 0
                            S = 0.1
                            W = 0.18
                            region = 'internal'
                        elif bl2 != bl3 and bl4 and bl5 and bl6:
                            L = 0
                            S = 0.0005
                            W = 0.97951
                            region = 'external'
                        elif X == 0 and Z == 0 and Y == trunkHeight:
                            L = 0
                            S = 0.0005
                            W = 0.97951
                            region = 'Xternal'
                        else:
                            L = 0
                            S = 0
                            W = 0
                            region = 'air2'

                        Block += 1

                        outSheet[bb].write(m + 1, 0, m)
                        outSheet[bb].write(m + 1, 1, X)
                        outSheet[bb].write(m + 1, 2, Y)
                        outSheet[bb].write(m + 1, 3, Z)
                        outSheet[bb].write(m + 1, 4, L * P)
                        outSheet[bb].write(m + 1, 5, S * P)
                        outSheet[bb].write(m + 1, 6, W * P)
                        outSheet[bb].write(m + 1, 7, P)
                        outSheet[bb].write(m + 1, 8, region)

                        m += 1
                        Z += 1
                    X += 1
                Y += 1

            # writing in headers
            outSheet[bb].write(0, 0, 'block index')
            outSheet[bb].write(0, 1, 'X offset')
            outSheet[bb].write(0, 2, 'Y offset')
            outSheet[bb].write(0, 3, 'Z offset')
            outSheet[bb].write(0, 4, 'stem chance')
            outSheet[bb].write(0, 5, 'shroom chance')
            outSheet[bb].write(0, 6, 'wart chance')
            outSheet[bb].write(0, 7, 'set chance')
            outSheet[bb].write(0, 8, 'region')

            print(f"LVS #{bb}: T: {trunkHeight} H: {hatHeight} F: {Fringe[N]} L: {LVS} P: {P}")
            bb += 1
            N += 1

print(Block)

outWorkbook.close()
