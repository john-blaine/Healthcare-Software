// These are four different methods for invoking a function
// which differ in how the parameter "this" is initialized

// This is the Method Invocation Pattern:
// Create myObject. It has a value and an increment
// method. The increment method takes an optional
// parameter. If the argument is not a number, then 1
// is used as the default.

var myObject = {
    value: 0,
    increment: function (inc) {
        this.value += typeof inc === 'number' ? inc : 1;
    },
    getValue: function () {
        return this.value;
    }
}

myObject.increment();
document.writeln(myObject.value);   // 1

myObject.increment(2);
document.writeln(myObject.value);   // 3

// This is the Function Invocation Pattern:
// When a function is not the property of an object,
// then it is invoked as a function:

var add = function (a, b) {
    return a + b;
};

var sum = add(3, 4);    // sum is 7
console.log(sum);

// Replace the "this" variable with the "self" variable so that
// the inner function's "this" (now "self") variable will reference
// the outer function object instead of the global object

myObject.double = function () {
    var self = this;    // Workaround.
    
    var helper = function () {
        self.value = add(self.value, self.value);
    };
    
    helper();   // Invoke helper as a function.
};

// Now we invoke Double as a method.

myObject.double();
document.writeln(myObject.getValue());  // 6

// This is the Constructor Invocation Pattern:

// Create a constructor function called Quo.
// It makes an object with a status property.

var Quo = function (string) {
    this.status = string;
};

// Give all instances of Quo a public method
// called get_status.

Quo.prototype.get_status = function () {
    return this.status;
};

// Make an instance of Quo.

var myQuo = new Quo('confused');

document.writeln(myQuo.get_status());

// This is the Apply Invocation Pattern:

// Make an array of 2 numbers and add them.

var array = [3, 4];
var sum = add.apply(null, array);   // sum is 7

// Make an object with a status member.

var statusObject = {
    status: 'A-OK'
};

test_2 = {
    subtract: function (a, b) {
        return a - b;
    }
};

test_1 = {};

test_2.subtract.apply(test_1, [1,2]);
console.log(test_2.subtract.apply(test_1, [5, 3]));


// statusObject does not inherit from Quo.prototype,
// but we can invoke the get_status method on
// statusObjects even though statusObjects does not
// have a get_status method.

var status = Quo.prototype.get_status.apply(statusObject);

// Here we are diverging to do some testing.

var Book = function boolean(worm_status) {
    this.worm = worm_status;
};

Book.prototype.worm_check = function () {
    return this.worm;
};

var myBook = new Book(true);

console.log(myBook.worm_check());

// Objects that are intended to be used with the "new" prefix are called constructor functions.
// The object below is a constructor. By using "this" and the dot notation, we can add methods to a constructor function.
// We can also use "Object".prototype."name of property or method to be added" to add that to the constructor
// prototype that will be used when objects are created from the constructor.

var Train = function boolean(wreck_status) {
    this.wreck = wreck_status
    this.wreck_check = function () {
        return this.wreck;
    }
};


var myTrain = new Train(false);

console.log(myTrain.wreck_check());
