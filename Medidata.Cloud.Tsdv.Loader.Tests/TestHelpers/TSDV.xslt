<?xml version="1.0"?>

<xsl:stylesheet version="1.0"
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns="urn:schemas-microsoft-com:office:spreadsheet"
                xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                xmlns:o="urn:schemas-microsoft-com:office:office"
                xmlns:x="urn:schemas-microsoft-com:office:excel"
                xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
  <xsl:output method="xml" indent="yes" encoding="utf-8" />
  <xsl:template match="/ArrayOfBlockPlan">
    <xsl:processing-instruction name="mso-application">
      <xsl:text>progid="Excel.Sheet"</xsl:text>
    </xsl:processing-instruction>


    <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
              xmlns:o="urn:schemas-microsoft-com:office:office"
              xmlns:x="urn:schemas-microsoft-com:office:excel"
              xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
              xmlns:html="http://www.w3.org/TR/REC-html40">
      <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
        <Author>Steve Mallory</Author>
        <LastAuthor>Steve Mallory</LastAuthor>
        <Version>14.00</Version>
      </DocumentProperties>
      <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
        <AllowPNG/>
      </OfficeDocumentSettings>
      <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
        <WindowHeight>11820</WindowHeight>
        <WindowWidth>18195</WindowWidth>
        <WindowTopX>480</WindowTopX>
        <WindowTopY>75</WindowTopY>
        <ProtectStructure>False</ProtectStructure>
        <ProtectWindows>False</ProtectWindows>
      </ExcelWorkbook>
      <Styles>
        <Style ss:ID="Default" ss:Name="Normal">
          <Alignment ss:Vertical="Bottom" />
          <Borders />
          <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000" />
          <Interior />
          <NumberFormat />
          <Protection />
        </Style>
        <Style ss:ID="sHeader">
          <Alignment ss:Vertical="Bottom" ss:WrapText="1" />
          <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="12" ss:Color="#000000" ss:Bold="1" />
        </Style>
        <Style ss:ID="sGeneral">
          <NumberFormat ss:Format="General" />
        </Style>
        <Style ss:ID="sGeneralNumber">
          <NumberFormat ss:Format="General Number" />
        </Style>
        <Style ss:ID="sGeneralDate">
          <NumberFormat ss:Format="General Date" />
        </Style>
        <Style ss:ID="sLongDate">
          <NumberFormat ss:Format="Long Date" />
        </Style>
        <Style ss:ID="sMediumDate">
          <NumberFormat ss:Format="Medium Date" />
        </Style>
        <Style ss:ID="sShortDate">
          <NumberFormat ss:Format="Short Date" />
        </Style>
        <Style ss:ID="sLongTime">
          <NumberFormat ss:Format="Long Time" />
        </Style>
        <Style ss:ID="sMediumTime">
          <NumberFormat ss:Format="Medium Time" />
        </Style>
        <Style ss:ID="sShortTime">
          <NumberFormat ss:Format="Short Time" />
        </Style>
        <Style ss:ID="sCurrency">
          <NumberFormat ss:Format="Currency" />
        </Style>
        <Style ss:ID="sEuroCurrency">
          <NumberFormat ss:Format="Euro Currency" />
        </Style>
        <Style ss:ID="sFixed">
          <NumberFormat ss:Format="Fixed" />
        </Style>
        <Style ss:ID="sStandard">
          <NumberFormat ss:Format="Standard" />
        </Style>
        <Style ss:ID="sPercent">
          <NumberFormat ss:Format="Percent" />
        </Style>
        <Style ss:ID="sScientific">
          <NumberFormat ss:Format="Scientific" />
        </Style>
        <Style ss:ID="sYes/No">
          <NumberFormat ss:Format="Yes/No" />
        </Style>
        <Style ss:ID="sTrue/False">
          <NumberFormat ss:Format="True/False" />
        </Style>
        <Style ss:ID="sOn/Off">
          <NumberFormat ss:Format="On/Off" />
        </Style>
      </Styles>
      <Worksheet ss:Name="BlockPlan">
        <Table>
          <xsl:for-each select="./BlockPlan">
            <Column ss:Width="150"/>
            <Column ss:Width="150"/>
            <Column ss:Width="150"/>
            <Column ss:Width="150"/>
            <Column ss:Width="150"/>
            <Column ss:Width="150"/>
            <Column ss:Width="150"/>
            <Column ss:Width="150"/>
            <Column ss:Width="150"/>
            <Column ss:Width="150"/>
            <Column ss:Width="150"/>
            <Row ss:AutoFitHeight="0">
            <Cell ss:StyleID="sHeader">
              <Data ss:Type="String">Block Plan Name</Data>
            </Cell>
              <Cell ss:StyleID="sHeader">
                <Data ss:Type="String">Block Plan Type</Data>
              </Cell>
              <Cell ss:StyleID="sHeader">
                <Data ss:Type="String">Study/SiteGroup/SiteName</Data>
              </Cell>
              <Cell ss:StyleID="sHeader">
                <Data ss:Type="String">Contains Subjects</Data>
              </Cell>
              <Cell ss:StyleID="sHeader">
                <Data ss:Type="String">Date Entry Role</Data>
              </Cell>
              <Cell ss:StyleID="sHeader">
                <Data ss:Type="String">Block Plan Status</Data>
              </Cell>
              <Cell ss:StyleID="sHeader">
                <Data ss:Type="String">Plan Activated By</Data>
              </Cell>
              <Cell ss:StyleID="sHeader">
                <Data ss:Type="String">Average Subjects/Site</Data>
              </Cell>
              <Cell ss:StyleID="sHeader">
                <Data ss:Type="String">Estimated Coverage</Data>
              </Cell>
              <Cell ss:StyleID="sHeader">
                <Data ss:Type="String">Using Matrix</Data>
              </Cell>
              <Cell ss:StyleID="sHeader">
                <Data ss:Type="String">Estimated Date</Data>
              </Cell>
              
              
            </Row>
            <Row ss:AutoFitHeight="0">
              <Cell ss:StyleID="sGeneral">
                <Data ss:Type="String">
                  <xsl:value-of select="Name" />
                </Data>
              </Cell>
              <Cell ss:StyleID="sGeneral">
                <Data ss:Type="String">
                  <xsl:value-of select="BlockPlanType" />
                </Data>
              </Cell>
              <Cell ss:StyleID="sGeneral">
                <Data ss:Type="String">
                  <xsl:value-of select="ObjectName" />
                </Data>
              </Cell>
              <Cell ss:StyleID="sYes/No">
                <Data ss:Type="Boolean">
                  <xsl:value-of select="number(not(IsProdInUse = 'false'))" />
                </Data>
              </Cell>
              <Cell ss:StyleID="sGeneral">
                <Data ss:Type="String">
                  <xsl:value-of select="RoleName" />
                </Data>
              </Cell>
              <Cell ss:StyleID="sYes/No">
                <Data ss:Type="Boolean">
                  <xsl:value-of select="number(not(Activated = 'false'))" />
                </Data>
              </Cell>
              <Cell ss:StyleID="sGeneral">
                <Data ss:Type="String">
                  <xsl:value-of select="ActivatedUserName" />
                </Data>
              </Cell>
              <Cell ss:StyleID="sGeneralNumber">
                <Data ss:Type="Number">
                  <xsl:value-of select="AverageSubjectPerSite" />
                </Data>
              </Cell>
              <Cell ss:StyleID="sGeneralNumber">
                <Data ss:Type="Number">
                  <xsl:value-of select="CoveragePercent" />
                </Data>
              </Cell>
              <Cell ss:StyleID="sGeneral">
                <Data ss:Type="String">
                  <xsl:value-of select="MatrixName" />
                </Data>
              </Cell>
              <Cell ss:StyleID="sGeneralDate">
                <Data ss:Type="DateTime">
                  <xsl:value-of select="substring(DateEstimated,0,20)" />
                </Data>
              </Cell>
            </Row>
          </xsl:for-each>
        </Table>
        <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
          <PageSetup>
            <Layout x:Orientation="Landscape"/>
            <Header x:Margin="0.3" x:Data="&amp;CTemplate"/>
            <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>
          </PageSetup>
          <Unsynced/>
          <FitToPage/>
          <Print>
            <ValidPrinterInfo/>
            <HorizontalResolution>600</HorizontalResolution>
            <VerticalResolution>600</VerticalResolution>
          </Print>
          <Selected/>
          <Panes>
            <Pane>
              <Number>3</Number>
              <ActiveRow>7</ActiveRow>
              <ActiveCol>2</ActiveCol>
            </Pane>
          </Panes>
          <ProtectObjects>False</ProtectObjects>
          <ProtectScenarios>False</ProtectScenarios>
        </WorksheetOptions>
      </Worksheet>
    </Workbook>

  </xsl:template>
</xsl:stylesheet>