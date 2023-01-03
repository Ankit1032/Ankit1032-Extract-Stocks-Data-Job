pipeline {
    agent {
        label 'windows'
    }
    stages {
        stage('Set property') {
            steps {
                bat 'System.setProperty("hudson.plugins.git.GitSCM.ALLOW_LOCAL_CHECKOUT", "TRUE")'
            }
        }
        stage('main') {
            steps {
                sh 'python main.py'
            }
        }
    }
}