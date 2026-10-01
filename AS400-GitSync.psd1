@{
    QueryScript        = 'D:\AS400\AS400-GitSync\SRC_MEMBER_LAST_CHANGE_1DAY.ps1'
    CsvPath            = 'D:\AS400\AS400-GitSync\SRC_MEMBER_LAST_CHANGE_1DAY.csv'
    CsvEncoding        = 'utf-8'

    IbmSystem          = 's105dd7a.dometic.internal'
    OdbcDriver         = 'IBM i Access ODBC Driver'
    CredentialPath     = 'D:\AS400\AS400-GitSync\ibmi-credential.xml'
    # Optional: System-DSN statt Driver/System benutzen. Muss zum selben IBM i zeigen!
    OdbcDsn            = ''
    # Optional: z.B. @{ SSL = '1' } bei eingerichtetem TLS.
    OdbcOptions        = @{}

    IfsRoot            = '/home/langlitz/AS400'
    # Windows-Zugriff auf GENAU IfsRoot (NetServer-Laufwerk oder UNC-Freigabe).
    IfsWindowsRoot     = 'F:\AS400'
    # Falls lokaler Clone: z.B. D:\Git\AS400; Dateien werden dann kopiert.
    RepoPath           = 'F:\AS400'
    Remote             = 'origin'
    Branch             = 'main'
    GitExe             = 'git.exe'

    LogDirectory       = 'D:\AS400\AS400-GitSync\Logs'
    QueryTimeoutSec    = 900
    CommandTimeoutSec  = 600
    GitTimeoutSec      = 300
}
