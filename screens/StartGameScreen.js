import React, { useState } from 'react';
import {
    View, StyleSheet, Text, Button, Alert,
    // 2 komponen dibawah ini diperlukan utk membuat keyboard menghilang saat user press di layar
    TouchableWithoutFeedback, Keyboard
} from 'react-native';

import Card from '../components/Card';
import Colors from '../constants/Colors';
import Input from '../components/Input';
import NumberContainer from '../components/NumberContainer';

const StartGameScreen = props => {

    const [enteredValue, setEnteredValue] = useState('');
    const [confirmed, setConfirmed] = useState(false);
    const [selectedNumber, setSelectedNumber] = useState();

    const numberInputhandler = inputText => {
        // regEx utk replace semua yang bukan 0-9 (bukan angka) menjadi string kosong
        setEnteredValue(inputText.replace(/[^0-9]/g, ''));
    };

    const resetInputHandler = () => {
        setEnteredValue('');
        setConfirmed(false);
    };

    const confirmInputHandler = () => {
        const chosenNumber = parseInt(enteredValue);
        if (isNaN(chosenNumber) || chosenNumber <= 0 || chosenNumber > 99) {
            Alert.alert('Invalid number!', 'Number has to be a number between 1 and 99.', [{ text: 'Okay', style: 'destructive', onPress: resetInputHandler }])
            return;
        }
        setConfirmed(true);
        setEnteredValue('');
        setSelectedNumber(chosenNumber);
        // Keyboard.dismiss() adalah fungsi untuk menghilangkan keyboard saat fungsi confirmInputHandler dijalankan
        Keyboard.dismiss();
    };

    let confirmedOutput;

    if (confirmed) {
        confirmedOutput = (
            <Card style={styles.summaryContainer} >
                <Text>You selected</Text>
                {/* jika tag nya seperti dibawah ini, maka selectedNumber bisa dipanggil di NumberContainer dengan props.children */}
                <NumberContainer>{selectedNumber}</NumberContainer>
                <Button title="START GAME" onPress={() => props.onStartGame(selectedNumber)} />
            </Card>
        );
    };

    return (
        // View dibungkus dengan TouchableWithoutFeedback agar bisa dikenakan onPress
        // Keyboard.dismiss() adalah fungsi untuk menghilangkan keyboard saat onPress diaktifkan
        <TouchableWithoutFeedback onPress={() => { Keyboard.dismiss() }} >
            <View style={styles.screen} >
                <Text style={styles.title} >Start a New Game!</Text>
                <Card style={styles.inputContainer} >
                    <Text>Select a Number</Text>
                    <Input
                        style={styles.input}
                        blurOnSubmit
                        autoCapitalize='none'
                        autoCorrect={false}
                        keyboardType='numeric'
                        maxLength={2}
                        onChangeText={numberInputhandler}
                        value={enteredValue}
                    />
                    <View style={styles.buttonContainer} >
                        <View style={styles.button} ><Button title="Reset" onPress={() => { resetInputHandler() }} color={Colors.accent} /></View>
                        <View style={styles.button} ><Button title="Confirm" onPress={() => { confirmInputHandler() }} color={Colors.primary} /></View>
                    </View>
                </Card>
                {confirmedOutput}
            </View>
        </TouchableWithoutFeedback>
    );
};

const styles = StyleSheet.create({
    screen: {
        flex: 1,
        padding: 10,
        alignItems: 'center'
    },
    title: {
        fontSize: 20,
        marginVertical: 10
    },
    inputContainer: {
        width: 300,
        maxWidth: '80%',
        alignItems: 'center'
    },
    buttonContainer: {
        flexDirection: 'row',
        width: '100%',
        justifyContent: 'space-between',
        paddingHorizontal: 15
    },
    button: {
        width: 100
    },
    input: {
        width: 50,
        textAlign: 'center'
    },
    summaryContainer: {
        marginTop: 20,
        // entah kenapa alignItems center di bawah ini juga membuat NumberContainer mengecil menjadi hanya sebesar angka nya
        alignItems: 'center'
    }
});

export default StartGameScreen;